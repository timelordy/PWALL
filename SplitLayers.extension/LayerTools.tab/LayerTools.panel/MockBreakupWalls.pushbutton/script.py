# -*- coding: utf-8 -*-
"""Разбивает составную стену на отдельные стены-слои, автоматически создавая тип для каждого слоя."""

import clr

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit.DB import (
    BuiltInParameter,
    CompoundStructure,
    CompoundStructureLayer,
    ElementId,
    FamilyInstance,
    FilteredElementCollector,
    LocationCurve,
    MaterialFunctionAssignment,
    Transform,
    UnitTypeId,
    UnitUtils,
    Wall,
    WallLocationLine,
    WallType,
    WallUtils,
    XYZ,
    JoinGeometryUtils,
)
from Autodesk.Revit.DB.Structure import StructuralType
from Autodesk.Revit.UI.Selection import ObjectType

from pyrevit import revit, forms, script

import re


doc = revit.doc
uidoc = revit.uidoc
logger = script.get_logger()

_WIDTH_EPS = 1e-6
_SIGNATURE_CACHE = {}


def _feet_to_mm(value):
    try:
        return UnitUtils.ConvertFromInternalUnits(value, UnitTypeId.Millimeters)
    except Exception:
        return value * 304.8


def _layers_to_sequence(layers):
    try:
        return list(layers) if layers is not None else []
    except TypeError:
        sequence = []
        try:
            count = getattr(layers, 'Count', None)
            if count is None:
                count = layers.Size
        except Exception:
            count = None
        if count is None:
            return sequence
        for idx in range(count):
            try:
                sequence.append(layers[idx])
            except Exception:
                try:
                    sequence.append(layers.get_Item(idx))
                except Exception:
                    break
        return sequence


def _collect_structure(wall):
    try:
        structure = wall.WallType.GetCompoundStructure()
    except Exception as exc:
        logger.debug('Не удалось получить CompoundStructure: %s', exc)
        return None, []

    if not structure:
        return None, []

    try:
        layers = list(structure.GetLayers())
    except Exception as exc:
        logger.debug('Не удалось получить слои: %s', exc)
        return structure, []

    return structure, layers


def _structure_layers_data(structure):
    layer_items = _layers_to_sequence(structure.GetLayers())
    first_core = getattr(structure, 'GetFirstCoreLayerIndex', lambda: -1)()
    last_core = getattr(structure, 'GetLastCoreLayerIndex', lambda: -1)()

    data = []
    position = 0.0
    for idx, layer in enumerate(layer_items):
        width = getattr(layer, 'Width', None) or 0.0
        start = position
        end = start + width
        position = end

        material_id = ElementId.InvalidElementId
        for attr in ('MaterialId', 'LayerMaterialId'):
            try:
                candidate = getattr(layer, attr)
            except Exception:
                candidate = None
            if candidate and isinstance(candidate, ElementId) and candidate.IntegerValue > 0:
                material_id = candidate
                break

        try:
            function = layer.Function
        except Exception:
            function = None

        is_core = (
            first_core is not None
            and last_core is not None
            and first_core != -1
            and last_core != -1
            and first_core <= idx <= last_core
        )

        data.append({
            'index': idx + 1,
            'width': width,
            'start': start,
            'end': end,
            'material_id': material_id,
            'function': function,
            'is_core': is_core,
        })

    total_width = position
    core_start = data[first_core]['start'] if data and first_core not in (-1, None) else 0.0
    core_end = data[last_core]['end'] if data and last_core not in (-1, None) else total_width

    return data, total_width, core_start, core_end


def _reference_offset(location_line_value, total_width, core_start, core_end):
    try:
        location_line = WallLocationLine(location_line_value)
    except Exception:
        location_line = WallLocationLine.WallCenterline

    if location_line in (
        WallLocationLine.WallCenterline,
        getattr(WallLocationLine, 'Centerline', WallLocationLine.WallCenterline),
        getattr(WallLocationLine, 'LineCenter', WallLocationLine.WallCenterline),
    ):
        return total_width / 2.0
    if location_line == WallLocationLine.FinishFaceExterior:
        return 0.0
    if location_line == WallLocationLine.FinishFaceInterior:
        return total_width
    if location_line == WallLocationLine.CoreFaceExterior:
        return core_start
    if location_line == WallLocationLine.CoreFaceInterior:
        return core_end
    if hasattr(WallLocationLine, 'CoreCenterline') and location_line == WallLocationLine.CoreCenterline:
        return (core_start + core_end) / 2.0
    if hasattr(WallLocationLine, 'CoreCenterlineExterior') and location_line == WallLocationLine.CoreCenterlineExterior:
        return (core_start + core_end) / 2.0
    if hasattr(WallLocationLine, 'CoreCenterlineInterior') and location_line == WallLocationLine.CoreCenterlineInterior:
        return (core_start + core_end) / 2.0
    return total_width / 2.0


def _compute_normal(curve):
    try:
        derivative = curve.ComputeDerivatives(0.5, True)
        tangent = derivative.BasisX.Normalize()
    except Exception:
        try:
            tangent = (curve.GetEndPoint(1) - curve.GetEndPoint(0)).Normalize()
        except Exception:
            return XYZ.BasisY
    normal = tangent.CrossProduct(XYZ.BasisZ)
    if normal.IsZeroLength():
        normal = XYZ.BasisY
    return normal.Normalize()


def _negate_vector(vector):
    if vector is None:
        return XYZ.Zero
    try:
        negated = vector.Negate()
        if negated:
            return negated
    except Exception:
        pass
    try:
        return XYZ(-getattr(vector, 'X', 0.0), -getattr(vector, 'Y', 0.0), -getattr(vector, 'Z', 0.0))
    except Exception:
        return XYZ.Zero


def _scale_vector(vector, scale):
    if vector is None:
        return XYZ.Zero
    try:
        return vector.Multiply(float(scale))
    except Exception:
        pass
    try:
        factor = float(scale)
    except Exception:
        factor = scale
    try:
        x = getattr(vector, 'X', 0.0)
        y = getattr(vector, 'Y', 0.0)
        z = getattr(vector, 'Z', 0.0)
        return XYZ(x * factor, y * factor, z * factor)
    except Exception:
        return XYZ.Zero


def _make_signature(layer_info):
    material_id = layer_info['material_id']
    material_value = material_id.IntegerValue if material_id and material_id.IntegerValue > 0 else -1

    function = layer_info['function']
    function_value = None
    if function is not None:
        try:
            function_value = int(function)
        except Exception:
            try:
                function_value = int(function.value__)
            except Exception:
                function_value = hash(str(function))
    else:
        function_value = -1

    return (round(layer_info['width'], 6), material_value, function_value)


def _find_existing_single_layer_type(signature):
    cached = _SIGNATURE_CACHE.get(signature)
    if cached:
        return cached

    width, material_value, function_value = signature

    collector = FilteredElementCollector(doc).OfClass(WallType)
    for candidate in collector:
        try:
            structure = candidate.GetCompoundStructure()
            layers = structure.GetLayers() if structure else None
        except Exception:
            continue
        layer_items = _layers_to_sequence(layers)
        if len(layer_items) != 1:
            continue

        layer = layer_items[0]
        width_candidate = getattr(layer, 'Width', None) or 0.0
        if abs(width_candidate - width) > _WIDTH_EPS:
            continue

        material_candidate = -1
        for attr in ('MaterialId', 'LayerMaterialId'):
            try:
                mat = getattr(layer, attr)
            except Exception:
                mat = None
            if mat and isinstance(mat, ElementId) and mat.IntegerValue > 0:
                material_candidate = mat.IntegerValue
                break

        if material_value != -1 and material_candidate != material_value:
            continue

        function_candidate = None
        try:
            function_candidate = layer.Function
        except Exception:
            pass

        if function_value != -1:
            match = False
            try:
                match = int(function_candidate) == function_value
            except Exception:
                try:
                    match = int(function_candidate.value__) == function_value
                except Exception:
                    match = hash(str(function_candidate)) == function_value
            if not match:
                continue

        _SIGNATURE_CACHE[signature] = candidate
        return candidate

    return None


def _sanitize_name(name):
    cleaned = re.sub(r'[\n\r<>:"/\\|?*]', '_', name)
    cleaned = re.sub(r'\s+', ' ', cleaned)
    return cleaned.strip()


def _clone_wall_type_for_layer(source_type, layer_info):
    signature = _make_signature(layer_info)
    existing = _find_existing_single_layer_type(signature)
    if existing:
        return existing

    width = layer_info['width']
    width_mm = _feet_to_mm(width)

    try:
        base_name = unicode(source_type.Name)
    except Exception:
        base_name = source_type.Name

    base_label = u"{} слой {} {:.1f} мм".format(base_name, layer_info['index'], width_mm)
    name = _sanitize_name(base_label)
    if not name:
        name = u"Layer_{}".format(layer_info['index'])

    suffix = 1
    while True:
        try:
            new_type = source_type.Duplicate(name)
            break
        except Exception:
            suffix += 1
            name = _sanitize_name(u"{} #{}".format(base_label, suffix))
            if suffix > 50:
                logger.warning('Не удалось дублировать тип для слоя %s', layer_info['index'])
                return None

    material_id = layer_info['material_id'] if layer_info['material_id'] else ElementId.InvalidElementId
    function = layer_info.get('function') or MaterialFunctionAssignment.Finish1

    try:
        new_layer = CompoundStructureLayer(width, function, material_id)
        if hasattr(new_layer, 'IsCore'):
            try:
                new_layer.IsCore = layer_info.get('is_core', False)
            except Exception:
                pass
        layer_list = List[CompoundStructureLayer]()
        layer_list.Add(new_layer)
        comp = CompoundStructure.CreateSimpleCompoundStructure(layer_list)
        new_type.SetCompoundStructure(comp)
    except Exception as exc:
        logger.warning('Не удалось пересобрать структуру типу "%s": %s', name, exc)
        return None

    _SIGNATURE_CACHE[signature] = new_type
    return new_type


def _apply_vertical_constraints(wall, context):
    try:
        param = wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET)
        if param is not None and context['base_offset'] is not None:
            param.Set(context['base_offset'])
    except Exception:
        pass
    try:
        top_level_id = context.get('top_level_id')
        if top_level_id and isinstance(top_level_id, ElementId) and top_level_id.IntegerValue > 0:
            param = wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE)
            if param:
                param.Set(top_level_id)
    except Exception:
        pass
    try:
        param = wall.get_Parameter(BuiltInParameter.WALL_TOP_OFFSET)
        if param is not None and context['top_offset'] is not None:
            param.Set(context['top_offset'])
    except Exception:
        pass


def _build_wall_context(wall, structure, total_width, core_start, core_end):
    location = getattr(wall, 'Location', None)
    if not isinstance(location, LocationCurve):
        raise ValueError(u'У стены отсутствует LocationCurve.')

    curve = location.Curve

    base_param = wall.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT)
    top_param = wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE)

    base_level_id = base_param.AsElementId() if base_param else ElementId.InvalidElementId
    top_level_id = top_param.AsElementId() if top_param else ElementId.InvalidElementId

    base_offset = wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET).AsDouble() if wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET) else 0.0
    top_offset = wall.get_Parameter(BuiltInParameter.WALL_TOP_OFFSET).AsDouble() if wall.get_Parameter(BuiltInParameter.WALL_TOP_OFFSET) else 0.0
    height = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).AsDouble() if wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM) else total_width

    flip = getattr(wall, 'Flipped', False)
    structural = getattr(wall, 'Structural', False)

    try:
        key_ref_param = wall.get_Parameter(BuiltInParameter.WALL_KEY_REF_PARAM)
        location_line = key_ref_param.AsInteger() if key_ref_param else int(WallLocationLine.WallCenterline)
    except Exception:
        location_line = int(WallLocationLine.WallCenterline)

    try:
        orientation = wall.Orientation
        if orientation and orientation.IsZeroLength():
            orientation = None
    except Exception:
        orientation = None

    reference_offset = _reference_offset(location_line, total_width, core_start, core_end)

    return {
        'curve': curve,
        'base_level_id': base_level_id if base_level_id and base_level_id.IntegerValue > 0 else wall.LevelId,
        'top_level_id': top_level_id if top_level_id and top_level_id.IntegerValue > 0 else None,
        'base_offset': base_offset,
        'top_offset': top_offset,
        'height': height,
        'flip': flip,
        'structural': structural,
        'location_line': location_line,
        'orientation': orientation,
        'reference_offset': reference_offset,
    }


def _ensure_orientation_vector(context, curve):
    orientation = context.get('orientation')
    if orientation and not orientation.IsZeroLength():
        return orientation.Normalize()
    return _compute_normal(curve)


def _collect_hosted_instances(wall):
    hosted = []
    fi_collector = FilteredElementCollector(doc).OfClass(FamilyInstance)
    for inst in fi_collector:
        host = getattr(inst, 'Host', None)
        if host is None or host.Id != wall.Id:
            continue

        location_point = None
        location_curve = None
        loc = inst.Location
        if isinstance(loc, LocationCurve):
            location_curve = loc.Curve
        else:
            try:
                location_point = loc.Point
            except Exception:
                location_point = None

        hosted.append({
            'id': inst.Id,
            'symbol': inst.Symbol,
            'level_id': inst.LevelId if inst.LevelId.IntegerValue > 0 else None,
            'location_point': location_point,
            'location_curve': location_curve,
            'hand_flipped': getattr(inst, 'HandFlipped', False),
            'face_flipped': getattr(inst, 'FacingFlipped', False),
        })
    return hosted


def _rehost_instances(instances, new_host_wall):
    for info in instances:
        symbol = info['symbol']
        if symbol is None:
            continue

        level = doc.GetElement(info['level_id']) if info['level_id'] else None
        if level is None:
            try:
                level = doc.GetElement(new_host_wall.LevelId)
            except Exception:
                level = None

        try:
            if info['location_curve'] is not None:
                new_inst = doc.Create.NewFamilyInstance(
                    info['location_curve'], symbol, new_host_wall, level, StructuralType.NonStructural
                )
            elif info['location_point'] is not None:
                new_inst = doc.Create.NewFamilyInstance(
                    info['location_point'], symbol, new_host_wall, level, StructuralType.NonStructural
                )
            else:
                continue
        except Exception as exc:
            logger.warning('Не удалось пересоздать хостируемый элемент %s: %s', info['id'], exc)
            continue

        try:
            doc.Delete(info['id'])
        except Exception:
            pass


def _breakup_wall(wall):
    structure, layers = _collect_structure(wall)
    if not layers:
        forms.alert(u'Выбранная стена не содержит слоёв.', exitscript=True)
        return

    layer_data, total_width, core_start, core_end = _structure_layers_data(structure)
    layer_data = [item for item in layer_data if item['width'] > _WIDTH_EPS]
    if not layer_data:
        forms.alert(u'Нет слоёв с ненулевой толщиной.', exitscript=True)
        return

    try:
        context = _build_wall_context(wall, structure, total_width, core_start, core_end)
    except ValueError as exc:
        forms.alert(unicode(exc), exitscript=True)
        return

    orientation = _ensure_orientation_vector(context, context['curve'])
    inward = _negate_vector(orientation)
    try:
        inward = inward.Normalize()
    except Exception:
        pass

    hosted_instances = _collect_hosted_instances(wall)

    base_curve = context['curve']
    base_level_id = context['base_level_id']

    created_walls = []
    t = Transaction(doc, 'Mock breakup wall into layers')
    t.Start()
    try:
        for layer_info in layer_data:
            layer_type = _clone_wall_type_for_layer(wall.WallType, layer_info)
            if layer_type is None:
                continue

            layer_center = (layer_info['start'] + layer_info['end']) / 2.0
            offset_center = layer_center - context['reference_offset']
            translation_vector = _scale_vector(inward, offset_center)
            placement_curve = base_curve.CreateTransformed(Transform.CreateTranslation(translation_vector))

            try:
                new_wall = Wall.Create(
                    doc,
                    placement_curve,
                    layer_type.Id,
                    base_level_id,
                    context['height'],
                    context['base_offset'],
                    context['flip'],
                    context['structural'],
                )
            except Exception as exc:
                logger.warning('Не удалось создать стену для слоя %s: %s', layer_info['index'], exc)
                continue

            _apply_vertical_constraints(new_wall, context)

            try:
                key_param = new_wall.get_Parameter(BuiltInParameter.WALL_KEY_REF_PARAM)
                if key_param:
                    key_param.Set(context['location_line'])
            except Exception:
                pass

            for end_idx in (0, 1):
                try:
                    WallUtils.AllowWallJoinAtEnd(new_wall, end_idx)
                except Exception:
                    pass

            created_walls.append(new_wall)
            logger.debug('Создана стеновая часть %s для слоя %s', new_wall.Id.IntegerValue, layer_info['index'])

        if not created_walls:
            logger.warning('Не удалось сформировать ни одной стены')
            t.RollBack()
            return

        structural_wall = None
        for wall_part, layer_info in zip(created_walls, layer_data):
            if layer_info.get('function') == MaterialFunctionAssignment.Structure:
                structural_wall = wall_part
                break
        if structural_wall is None:
            structural_wall = created_walls[0]

        _rehost_instances(hosted_instances, structural_wall)

        for i in range(len(created_walls)):
            for j in range(i + 1, len(created_walls)):
                try:
                    JoinGeometryUtils.JoinGeometry(doc, created_walls[i], created_walls[j])
                except Exception:
                    pass

        try:
            doc.Delete(wall.Id)
        except Exception as exc:
            logger.warning('Не удалось удалить исходную стену: %s', exc)
            t.RollBack()
            return

        t.Commit()
        forms.alert(u'Создано {} стен. Проверьте соединения.'.format(len(created_walls)))
    except Exception as exc:
        logger.exception('Ошибка выполнения MockBreakupWalls: %s', exc)
        t.RollBack()
        forms.alert(u'Ошибка: {}'.format(exc), exitscript=True)
        return


def main():
    try:
        reference = uidoc.Selection.PickObject(
            ObjectType.Element,
            u"Выберите составную стену для разбивки"
        )
    except Exception:
        return

    wall = doc.GetElement(reference.ElementId)
    if not isinstance(wall, Wall):
        forms.alert(u'Выбранный элемент не является стеной.', exitscript=True)
        return

    _breakup_wall(wall)


if __name__ == '__main__':
    main()
