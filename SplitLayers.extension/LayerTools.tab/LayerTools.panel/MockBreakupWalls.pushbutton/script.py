# -*- coding: utf-8 -*-
"""Разбивает составную стену на отдельные стены-слои, автоматически создавая тип для каждого слоя."""

import clr

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System')
clr.AddReference('System.Xml')
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')

from System.Collections.Generic import List
from collections import defaultdict
from System.IO import StringReader
from System.Xml import XmlReader
from System.Windows import SystemParameters, WindowStartupLocation
from System.Windows.Input import Key
from System.Windows.Markup import XamlReader

from Autodesk.Revit.DB import (
    BuiltInParameter,
    CompoundStructure,
    CompoundStructureLayer,
    ElementId,
    FamilyInstance,
    FilteredElementCollector,
    ElementClassFilter,
    LocationCurve,
    LocationPoint,
    Opening,
    MaterialFunctionAssignment,
    Transform,
    Transaction,
    UnitTypeId,
    UnitUtils,
    Wall,
    WallLocationLine,
    WallType,
    WallUtils,
    Line,
    XYZ,
    JoinGeometryUtils,
    InstanceVoidCutUtils,
    GeometryInstance,
    Options,
    Solid,
    ViewDetailLevel,
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
_LAYER_JOIN_CACHE = {}
_PENDING_LAYER_JOINS = defaultdict(list)
_OPENING_MARGIN = 0.0
_OFFSET_TOLERANCE = 1e-4

try:
    _REVIT_MIN_DIMENSION = UnitUtils.ConvertToInternalUnits(1.0 / 32.0, UnitTypeId.Feet)
except Exception:
    _REVIT_MIN_DIMENSION = 1.0 / 32.0

_CAN_ADD_VOID_CUT = getattr(InstanceVoidCutUtils, 'CanAddInstanceVoidCut', None)
_ADD_INSTANCE_VOID_CUT = getattr(InstanceVoidCutUtils, 'AddInstanceVoidCut', None)
_APPLY_WALL_JOIN_TYPE = getattr(WallUtils, 'ApplyJoinType', None)
_ALLOW_WALL_JOIN_AT_END = getattr(WallUtils, 'AllowWallJoinAtEnd', None)
_DISALLOW_WALL_JOIN_AT_END = getattr(WallUtils, 'DisallowWallJoinAtEnd', None)

try:
    from Autodesk.Revit.DB import WallJoinType  # type: ignore
except Exception:
    WallJoinType = None

try:
    _WALL_JOIN_TYPE_BUTT = WallJoinType.Butt if WallJoinType is not None else None
except Exception:
    _WALL_JOIN_TYPE_BUTT = None

try:
    _TEXT_TYPE = unicode
except NameError:
    _TEXT_TYPE = str


def _to_unicode(value):
    if value is None:
        return u''
    try:
        return _TEXT_TYPE(value)
    except Exception:
        try:
            return _TEXT_TYPE(str(value))
        except Exception:
            return u''


_LAYER_FUNCTION_MAP = {
    MaterialFunctionAssignment.Structure: u"Несущий слой",
    MaterialFunctionAssignment.Substrate: u"Основание",
    MaterialFunctionAssignment.Insulation: u"Утеплитель",
    MaterialFunctionAssignment.Finish1: u"Отделка (наружная)",
    MaterialFunctionAssignment.Finish2: u"Отделка (внутренняя)",
    MaterialFunctionAssignment.Membrane: u"Мембрана",
}

_DEFAULT_LAYER_FUNCTION = getattr(MaterialFunctionAssignment, 'Other', None)
if _DEFAULT_LAYER_FUNCTION is not None and _DEFAULT_LAYER_FUNCTION not in _LAYER_FUNCTION_MAP:
    _LAYER_FUNCTION_MAP[_DEFAULT_LAYER_FUNCTION] = u"Прочий слой"


def _describe_layer_function(layer_function):
    try:
        return _LAYER_FUNCTION_MAP[layer_function]
    except Exception:
        return _to_unicode(layer_function)


def _get_param_double(element, param_id):
    if element is None:
        return None
    try:
        param = element.get_Parameter(param_id)
    except Exception:
        param = None
    if param is None:
        return None
    try:
        if hasattr(param, 'HasValue') and not param.HasValue:
            return None
    except Exception:
        pass
    try:
        return param.AsDouble()
    except Exception:
        try:
            return float(param.AsValueString())
        except Exception:
            return None


def _get_instance_clear_width(instance):
    if instance is None:
        return None

    builtin_names = (
        'DOOR_WIDTH',
        'WINDOW_WIDTH',
        'INSTANCE_WIDTH_PARAM',
        'FAMILY_WIDTH_PARAM',
        'SYMBOL_WIDTH_PARAM',
        'DOOR_PANEL_WIDTH',
        'DOOR_OPENING_WIDTH',
    )
    for param_name in builtin_names:
        param_id = getattr(BuiltInParameter, param_name, None)
        if param_id is None:
            continue
        value = _get_param_double(instance, param_id)
        if value is not None and value > _WIDTH_EPS:
            return value

    fallback_names = (
        u'Clear Width',
        u'Opening Width',
        u'Rough Width',
        u'Width',
        u'WIDTH',
        u'Nominal Width',
        u'Номинальная ширина',
        u'Ширина',
        u'ширина',
    )

    for name in fallback_names:
        try:
            param = instance.LookupParameter(name)
        except Exception:
            param = None
        value = None
        if param is not None:
            try:
                value = param.AsDouble()
            except Exception:
                value = None
        if value is not None and value > _WIDTH_EPS:
            return value

    symbol = getattr(instance, 'Symbol', None)
    if symbol is not None:
        for name in fallback_names:
            try:
                param = symbol.LookupParameter(name)
            except Exception:
                param = None
            value = None
            if param is not None:
                try:
                    value = param.AsDouble()
                except Exception:
                    value = None
            if value is not None and value > _WIDTH_EPS:
                return value

    return None



def _collect_geometry_edge_points(geometry, transform=None):
    if transform is None:
        transform = Transform.Identity

    points = []
    if geometry is None:
        return points

    for obj in geometry:
        if isinstance(obj, Solid):
            skip_solid = False
            try:
                if getattr(obj, 'Volume', None) is not None and obj.Volume <= _WIDTH_EPS:
                    skip_solid = True
            except Exception:
                pass
            if skip_solid:
                continue

            collected = False
            try:
                edges = obj.Edges
            except Exception:
                edges = None
            if edges is not None:
                try:
                    iterator = iter(edges)
                except TypeError:
                    iterator = None
                else:
                    for edge in edges:
                        try:
                            curve = edge.AsCurve()
                        except Exception:
                            curve = None
                        if curve is None:
                            continue
                        for end_idx in (0, 1):
                            try:
                                pt = curve.GetEndPoint(end_idx)
                            except Exception:
                                pt = None
                            if pt is None:
                                continue
                            try:
                                points.append(transform.OfPoint(pt))
                                collected = True
                            except Exception:
                                pass
            if not collected:
                try:
                    bbox = obj.GetBoundingBox()
                except Exception:
                    bbox = None
                if bbox is not None:
                    raw_points = [
                        XYZ(bbox.Min.X, bbox.Min.Y, bbox.Min.Z),
                        XYZ(bbox.Max.X, bbox.Min.Y, bbox.Min.Z),
                        XYZ(bbox.Min.X, bbox.Max.Y, bbox.Min.Z),
                        XYZ(bbox.Max.X, bbox.Max.Y, bbox.Min.Z),
                        XYZ(bbox.Min.X, bbox.Min.Y, bbox.Max.Z),
                        XYZ(bbox.Max.X, bbox.Min.Y, bbox.Max.Z),
                        XYZ(bbox.Min.X, bbox.Max.Y, bbox.Max.Z),
                        XYZ(bbox.Max.X, bbox.Max.Y, bbox.Max.Z),
                    ]
                    for pt in raw_points:
                        try:
                            points.append(transform.OfPoint(pt))
                        except Exception:
                            pass
        elif isinstance(obj, GeometryInstance):
            nested_transform = transform.Multiply(obj.Transform)
            try:
                nested_geometry = obj.GetSymbolGeometry()
            except Exception:
                try:
                    nested_geometry = obj.GetInstanceGeometry()
                except Exception:
                    nested_geometry = None
            nested_points = _collect_geometry_edge_points(nested_geometry, nested_transform)
            if nested_points:
                points.extend(nested_points)

    return points


def _get_instance_geometry_metrics(instance, width_direction=None):
    if instance is None:
        return {}

    try:
        options = Options()
        options.DetailLevel = ViewDetailLevel.Fine
        options.IncludeNonVisibleObjects = False
    except Exception:
        options = None

    geometry = None
    try:
        geometry = instance.get_Geometry(options) if options is not None else instance.get_Geometry(None)
    except Exception:
        try:
            geometry = instance.get_Geometry(None)
        except Exception:
            geometry = None

    points = _collect_geometry_edge_points(geometry)
    if not points:
        return {}

    metrics = {}

    if width_direction is not None and not getattr(width_direction, 'IsZeroLength', lambda: True)():
        try:
            direction = width_direction.Normalize()
        except Exception:
            direction = None
        if direction is not None:
            projections = [direction.DotProduct(pt) for pt in points]
            if projections:
                metrics['width'] = max(projections) - min(projections)

    heights = [pt.Z for pt in points]
    if heights:
        metrics['bottom'] = min(heights)
        metrics['top'] = max(heights)
        metrics['height'] = metrics['top'] - metrics['bottom']

    return metrics


def _element_id_to_int(elem_id):
    if isinstance(elem_id, ElementId):
        return elem_id.IntegerValue
    try:
        return int(elem_id) if elem_id is not None else None
    except Exception:
        return None


def _layer_signature_for_join(layer_info):
    material_id = layer_info.get('material_id')
    material_key = _element_id_to_int(material_id)
    width = round(layer_info.get('width', 0.0), 6)
    function = layer_info.get('function')
    return (material_key, width, function)


def _get_layer_offset(layer):
    if not isinstance(layer, dict):
        return None
    offset = layer.get('offset')
    if offset is None:
        start = layer.get('start')
        end = layer.get('end')
        if start is None or end is None:
            return None
        try:
            offset = (float(start) + float(end)) / 2.0
        except Exception:
            return None
        reference_offset = layer.get('reference_offset')
        if reference_offset is not None:
            try:
                offset -= float(reference_offset)
            except Exception:
                pass
    try:
        return float(offset)
    except Exception:
        return None


def _offset_difference(layer_a, layer_b):
    offset_a = _get_layer_offset(layer_a)
    offset_b = _get_layer_offset(layer_b)
    if offset_a is None or offset_b is None:
        return None
    try:
        return abs(offset_a - offset_b)
    except Exception:
        return None


def _offsets_compatible(layer_a, layer_b, tolerance=_OFFSET_TOLERANCE):
    diff = _offset_difference(layer_a, layer_b)
    if diff is None:
        return True
    return diff <= tolerance


def _match_layer_by_signature(target_layer, candidates, used_ids):
    target_signature = _layer_signature_for_join(target_layer)
    target_offset = _get_layer_offset(target_layer)
    best_candidate = None
    best_diff = None

    for candidate in candidates:
        candidate_id = _element_id_to_int(candidate.get('wall_id'))
        if candidate_id in used_ids:
            continue
        if _layer_signature_for_join(candidate) != target_signature:
            continue

        if target_offset is None:
            return candidate

        diff = _offset_difference(target_layer, candidate)
        if diff is not None and diff <= _OFFSET_TOLERANCE:
            return candidate

        if diff is None:
            if best_candidate is None:
                best_candidate = candidate
            continue

        if best_diff is None or diff < best_diff:
            best_candidate = candidate
            best_diff = diff

    return best_candidate


def _normalize_join_meta(meta):
    result = {'self_cuts': None, 'allow_mismatch': False}
    if isinstance(meta, dict):
        if 'self_cuts' in meta:
            result['self_cuts'] = meta.get('self_cuts')
        if meta.get('allow_mismatch'):
            result['allow_mismatch'] = True
    return result


def _entry_identity(entry):
    if not isinstance(entry, dict):
        return None

    wall_type_id = entry.get('wall_type_id')
    try:
        wall_type_key = int(wall_type_id) if wall_type_id is not None else None
    except Exception:
        wall_type_key = _element_id_to_int(wall_type_id)

    layer_ids = []
    for layer in entry.get('layers') or []:
        layer_ids.append(_element_id_to_int((layer or {}).get('wall_id')))

    if not layer_ids:
        return None

    return (wall_type_key, tuple(sorted(layer_ids)))


def _enqueue_pending_join(target_wall_id, entry, meta):
    if target_wall_id is None or not entry:
        return

    normalized_meta = _normalize_join_meta(meta)
    identity = _entry_identity(entry)
    if identity is None:
        return

    pending_items = _PENDING_LAYER_JOINS[target_wall_id]

    for item in pending_items:
        if isinstance(item, dict):
            existing_entry = item.get('entry')
            existing_meta = _normalize_join_meta(item.get('meta'))
            existing_identity = item.get('identity')
        elif isinstance(item, (list, tuple)) and len(item) == 2:
            existing_entry, existing_meta = item
            existing_identity = _entry_identity(existing_entry)
            existing_meta = _normalize_join_meta(existing_meta)
        else:
            existing_entry = item
            existing_meta = _normalize_join_meta(None)
            existing_identity = _entry_identity(existing_entry)

        if existing_identity == identity and existing_meta == normalized_meta:
            return

    pending_items.append({'entry': entry, 'meta': normalized_meta, 'identity': identity})


def _build_join_entry_from_existing_wall(wall, expected_signatures=None, target_type_id=None):
    if not isinstance(wall, Wall):
        return None

    structure, layers = _collect_structure(wall)
    if not layers:
        return None

    layer_data, total_width, core_start, core_end = _structure_layers_data(structure)
    try:
        context = _build_wall_context(wall, structure, total_width, core_start, core_end)
    except Exception:
        context = {}
    reference_offset = context.get('reference_offset')
    if reference_offset is None:
        reference_offset = total_width / 2.0 if total_width is not None else None
    filtered_layers = []
    for info in layer_data:
        width = info.get('width', 0.0) or 0.0
        if width <= _WIDTH_EPS:
            continue
        if expected_signatures:
            signature = _layer_signature_for_join(info)
            if signature not in expected_signatures:
                continue
        start = info.get('start')
        end = info.get('end')
        center = None
        if start is not None and end is not None:
            try:
                center = (float(start) + float(end)) / 2.0
            except Exception:
                center = None
        offset = None
        if center is not None and reference_offset is not None:
            try:
                offset = center - float(reference_offset)
            except Exception:
                offset = None
        filtered_layers.append({
            'index': info.get('index'),
            'function': info.get('function'),
            'width': width,
            'material_id': info.get('material_id'),
            'wall_id': wall.Id,
            'offset': offset,
            'reference_offset': reference_offset,
        })

    if not filtered_layers:
        return None

    actual_type_id = _element_id_to_int(getattr(wall.WallType, 'Id', None))

    entry = {
        'wall_type_id': target_type_id if target_type_id is not None else actual_type_id,
        'layers': filtered_layers,
    }

    if actual_type_id is not None:
        entry['actual_wall_type_id'] = actual_type_id

    return entry


def _apply_butt_join_type(wall):
    if wall is None:
        return
    if _APPLY_WALL_JOIN_TYPE is None or _WALL_JOIN_TYPE_BUTT is None:
        return
    for end_idx in (0, 1):
        try:
            _APPLY_WALL_JOIN_TYPE(wall, end_idx, _WALL_JOIN_TYPE_BUTT)
        except Exception:
            continue


def _ensure_join_controls(wall):
    if wall is None or _ALLOW_WALL_JOIN_AT_END is None:
        return
    for end_idx in (0, 1):
        try:
            _ALLOW_WALL_JOIN_AT_END(wall, end_idx)
        except Exception:
            continue


def _suppress_auto_join(wall):
    if wall is None or _DISALLOW_WALL_JOIN_AT_END is None:
        return False
    success = False
    for end_idx in (0, 1):
        try:
            _DISALLOW_WALL_JOIN_AT_END(wall, end_idx)
            success = True
        except Exception:
            continue
    return success


def _unjoin_two_walls(first_wall, second_wall):
    if first_wall is None or second_wall is None:
        return False
    first_id = getattr(getattr(first_wall, 'Id', None), 'IntegerValue', None)
    second_id = getattr(getattr(second_wall, 'Id', None), 'IntegerValue', None)
    if first_id is not None and second_id is not None and first_id == second_id:
        return False
    try:
        JoinGeometryUtils.UnjoinGeometry(doc, first_wall, second_wall)
        return True
    except Exception:
        return False


def _prepare_wall_for_manual_join(wall, existing_walls=None, host_wall=None):
    if wall is None:
        return

    _ensure_join_controls(wall)
    _suppress_auto_join(wall)

    neighbors = []
    if existing_walls:
        for candidate in existing_walls:
            if candidate is None:
                continue
            cand_id = getattr(getattr(candidate, 'Id', None), 'IntegerValue', None)
            wall_id = getattr(getattr(wall, 'Id', None), 'IntegerValue', None)
            if cand_id is not None and wall_id is not None and cand_id == wall_id:
                continue
            neighbors.append(candidate)

    if host_wall is not None:
        host_id = getattr(getattr(host_wall, 'Id', None), 'IntegerValue', None)
        wall_id = getattr(getattr(wall, 'Id', None), 'IntegerValue', None)
        if host_id is None or wall_id is None or host_id != wall_id:
            neighbors.append(host_wall)

    if neighbors:
        try:
            doc.Regenerate()
        except Exception:
            pass

    for neighbor in neighbors:
        _unjoin_two_walls(wall, neighbor)


def _join_two_walls(first_id, second_id, should_first_cut=None):
    if first_id is None or second_id is None:
        return False
    try:
        first_wall = doc.GetElement(first_id)
        second_wall = doc.GetElement(second_id)
    except Exception:
        return False
    if first_wall is None or second_wall is None:
        return False

    try:
        doc.Regenerate()
    except Exception:
        pass

    for end_idx in (0, 1):
        for wall_obj in (first_wall, second_wall):
            try:
                WallUtils.AllowWallJoinAtEnd(wall_obj, end_idx)
            except Exception:
                pass

    try:
        JoinGeometryUtils.JoinGeometry(doc, first_wall, second_wall)
    except Exception:
        pass

    if not JoinGeometryUtils.AreElementsJoined(doc, first_wall, second_wall):
        return False

    if should_first_cut is not None:
        try:
            is_cutting = JoinGeometryUtils.IsCuttingElementInJoin(doc, first_wall, second_wall)
        except Exception:
            is_cutting = None
        if is_cutting is not None and is_cutting != should_first_cut:
            try:
                JoinGeometryUtils.SwitchJoinOrder(doc, first_wall, second_wall)
            except Exception:
                pass

    try:
        doc.Regenerate()
    except Exception:
        pass

    _apply_butt_join_type(first_wall)
    _apply_butt_join_type(second_wall)

    return True


def _attempt_layer_joins(entry_a, entry_b, join_meta=None):
    if not entry_a or not entry_b:
        return False

    type_a = entry_a.get('wall_type_id')
    type_b = entry_b.get('wall_type_id')
    allow_mismatch = False
    if isinstance(join_meta, dict):
        allow_mismatch = bool(join_meta.get('allow_mismatch'))
    if (type_a is None or type_b is None or type_a != type_b) and not allow_mismatch:
        return False

    should_first_cut = None
    if isinstance(join_meta, dict):
        should_first_cut = join_meta.get('self_cuts')

    layers_a = entry_a.get('layers') or []
    layers_b = entry_b.get('layers') or []
    if not layers_a or not layers_b:
        return False

    layers_b_by_index = {layer.get('index'): layer for layer in layers_b}
    used_b = set()
    joined_any = False

    for layer_a in layers_a:
        layer_b = layers_b_by_index.get(layer_a.get('index'))
        if layer_b is None or _element_id_to_int(layer_b.get('wall_id')) in used_b:
            layer_b = _match_layer_by_signature(layer_a, layers_b, used_b)
        elif not _offsets_compatible(layer_a, layer_b):
            temp_used = set(used_b)
            candidate_id = _element_id_to_int(layer_b.get('wall_id'))
            if candidate_id is not None:
                temp_used.add(candidate_id)
            layer_b = _match_layer_by_signature(layer_a, layers_b, temp_used)
        if layer_b is None:
            continue

        if _join_two_walls(layer_a.get('wall_id'), layer_b.get('wall_id'), should_first_cut):
            used_b.add(_element_id_to_int(layer_b.get('wall_id')))
            joined_any = True

    return joined_any



def _invert_join_meta(join_meta):
    result = {'self_cuts': None, 'allow_mismatch': False}
    if not isinstance(join_meta, dict):
        return result
    value = join_meta.get('self_cuts') if join_meta else None
    if value is not None:
        result['self_cuts'] = not value
    if join_meta.get('allow_mismatch'):
        result['allow_mismatch'] = True
    return result


def _collect_joined_wall_ids(wall, wall_type_id, layer_data=None):
    results = []
    if wall is None:
        return results

    expected_signatures = None
    if layer_data:
        expected_signatures = set()
        for info in layer_data:
            try:
                expected_signatures.add(_layer_signature_for_join(info))
            except Exception:
                continue

    try:
        joined_elements = JoinGeometryUtils.GetJoinedElements(doc, wall)
    except Exception:
        joined_elements = []

    unique_ids = set()
    for elem_id in joined_elements or []:
        key = _element_id_to_int(elem_id)
        if key is None or key == wall.Id.IntegerValue:
            continue
        if key in unique_ids:
            continue

        neighbour = doc.GetElement(elem_id)
        if neighbour is None:
            continue

        neighbour_type = None
        try:
            neighbour_type = _element_id_to_int(neighbour.WallType.Id)
        except Exception:
            neighbour_type = None

        info = {
            'id': key,
            'self_cuts': None,
            'entry': None,
            'allow_mismatch': False,
        }

        try:
            info['self_cuts'] = JoinGeometryUtils.IsCuttingElementInJoin(doc, wall, neighbour)
        except Exception:
            pass

        cached_entry = _LAYER_JOIN_CACHE.get(key)
        if cached_entry is not None:
            if wall_type_id is None or cached_entry.get('wall_type_id') in (None, wall_type_id):
                info['entry'] = cached_entry
                results.append(info)
                unique_ids.add(key)
            continue

        if isinstance(neighbour, Wall):
            if wall_type_id is None or neighbour_type == wall_type_id:
                results.append(info)
                unique_ids.add(key)
                continue

            if expected_signatures:
                entry = _build_join_entry_from_existing_wall(
                    neighbour,
                    expected_signatures=expected_signatures,
                    target_type_id=wall_type_id,
                )
                if entry and entry.get('layers'):
                    info['entry'] = entry
                    info['allow_mismatch'] = True
                    _LAYER_JOIN_CACHE[key] = entry
                    results.append(info)
                    unique_ids.add(key)

    return results


def _normalize_join_info(join_info):
    if isinstance(join_info, dict):
        result = dict(join_info)
        result['id'] = _element_id_to_int(result.get('id'))
        result.setdefault('self_cuts', None)
        if 'entry' not in result:
            result['entry'] = None
        result['allow_mismatch'] = bool(result.get('allow_mismatch'))
        return result
    return {
        'id': _element_id_to_int(join_info),
        'self_cuts': None,
        'entry': None,
        'allow_mismatch': False,
    }


def _handle_layer_joins(original_wall_id, wall_type_id, produced_layers, joined_wall_ids):
    if not produced_layers:
        return

    layer_records = []
    for record in produced_layers:
        info = record.get('info') or {}
        wall = record.get('wall')
        if wall is None:
            continue
        offset = record.get('offset')
        if offset is None:
            offset = info.get('offset')
        reference_offset = record.get('reference_offset')
        if reference_offset is None:
            reference_offset = info.get('reference_offset')
        layer_records.append({
            'index': info.get('index'),
            'function': info.get('function'),
            'width': info.get('width', 0.0),
            'material_id': info.get('material_id'),
            'wall_id': wall.Id,
            'offset': offset,
            'reference_offset': reference_offset,
        })

    if not layer_records:
        return

    entry = {
        'wall_type_id': wall_type_id,
        'layers': layer_records,
    }

    _LAYER_JOIN_CACHE[original_wall_id] = entry

    pending_entries = _PENDING_LAYER_JOINS.pop(original_wall_id, [])
    for pending_item in pending_entries:
        if isinstance(pending_item, dict):
            pending_entry = pending_item.get('entry')
            pending_meta = pending_item.get('meta')
        elif isinstance(pending_item, (list, tuple)) and len(pending_item) == 2:
            pending_entry, pending_meta = pending_item
        else:
            pending_entry = pending_item
            pending_meta = {'self_cuts': None}
        _attempt_layer_joins(entry, pending_entry, pending_meta)

    for neighbour_info in joined_wall_ids or []:
        info = _normalize_join_info(neighbour_info)
        neighbour_id = info.get('id')
        if neighbour_id is None:
            continue

        neighbour_entry = info.get('entry')
        if not neighbour_entry:
            neighbour_entry = _LAYER_JOIN_CACHE.get(neighbour_id)
        if neighbour_entry:
            _attempt_layer_joins(entry, neighbour_entry, info)
        _enqueue_pending_join(neighbour_id, entry, _invert_join_meta(info))


class _LayerChoice(object):
    def __init__(self, layer_info):
        self.layer_info = layer_info
        function_name = _describe_layer_function(layer_info.get('function'))
        material_name = _to_unicode(layer_info.get('material_name') or u'')
        width_mm = _feet_to_mm(layer_info.get('width', 0.0))
        try:
            width_text = u"{:.1f}".format(width_mm)
        except Exception:
            width_text = _to_unicode(width_mm)
        details = []
        if material_name:
            details.append(u"Материал: {}".format(material_name))
        details.append(u"{} мм".format(width_text))
        detail_text = u", ".join(details)
        core_suffix = u" (ядро)" if layer_info.get('is_core') else u""
        self.name = u"Слой {index}{core}: {function} — {details}".format(
            index=layer_info.get('index'),
            core=core_suffix,
            function=function_name,
            details=detail_text,
        )


_LAYER_SELECTION_XAML = u"""
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Выберите слой, который останется хостом размещенных элементов"
        Width="360"
        Height="380"
        WindowStartupLocation="Manual"
        ResizeMode="CanResizeWithGrip"
        ShowInTaskbar="False"
        Topmost="True"
        AllowsTransparency="False">
    <DockPanel Margin="12">
        <Button x:Name="OkButton"
                Content="Выбрать"
                DockPanel.Dock="Bottom"
                Width="110"
                Height="28"
                HorizontalAlignment="Right"
                Margin="0,12,0,0"/>
        <ListBox x:Name="LayerList"
                 DisplayMemberPath="name"
                 HorizontalContentAlignment="Stretch"
                 ScrollViewer.VerticalScrollBarVisibility="Auto"/>
    </DockPanel>
</Window>
"""


class _LayerSelectionDialog(object):
    def __init__(self, choices, default_choice=None):
        self._choices = choices or []
        reader = XmlReader.Create(StringReader(_LAYER_SELECTION_XAML))
        self._window = XamlReader.Load(reader)
        self._window.WindowStartupLocation = WindowStartupLocation.Manual
        self._listbox = self._window.FindName('LayerList')
        self._ok_button = self._window.FindName('OkButton')
        self._listbox.ItemsSource = self._choices
        if default_choice in self._choices:
            self._listbox.SelectedItem = default_choice
        elif self._choices:
            self._listbox.SelectedIndex = 0
        self._ok_button.Click += self._on_accept
        self._window.KeyDown += self._on_key_down
        self._listbox.MouseDoubleClick += self._on_double_click
        self._result = None

    def _position_window(self):
        try:
            work_area = SystemParameters.WorkArea
            width = self._window.Width if self._window.Width > 0 else 360
            height = self._window.Height if self._window.Height > 0 else 380
            self._window.Left = max(work_area.Left, work_area.Right - width - 60)
            self._window.Top = max(work_area.Top, work_area.Top + 60)
        except Exception:
            pass

    def _on_key_down(self, sender, args):
        try:
            key = args.Key
        except Exception:
            key = None
        if key == Key.Enter:
            self._accept()
        elif key == Key.Escape:
            self._result = None
            self._window.Close()

    def _on_double_click(self, sender, args):
        self._accept()

    def _on_accept(self, sender, args):
        self._accept()

    def _accept(self):
        try:
            self._result = self._listbox.SelectedItem
        except Exception:
            self._result = None
        self._window.Close()

    def show_dialog(self):
        self._position_window()
        self._window.ShowDialog()
        return self._result

    def __unicode__(self):
        return self.name

    def __str__(self):  # pragma: no cover - совместимость Python 2/3
        try:
            return self.name.encode('utf-8')
        except Exception:
            return _to_unicode(self.name)


def _get_element_name(element):
    if element is None:
        return u''
    for attr in ('Name', 'name'):
        try:
            candidate = getattr(element, attr)
            if callable(candidate):
                candidate = candidate()
            if candidate not in (None, ''):
                return _to_unicode(candidate)
        except Exception:
            continue
    try:
        param = element.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
        if param:
            candidate = param.AsString()
            if candidate:
                return _to_unicode(candidate)
    except Exception:
        pass
    try:
        return _to_unicode(element)
    except Exception:
        return u''





def _adjust_view_detail_for_preview():
    state = {'param': None, 'previous': None, 'changed': False}
    try:
        active_view = uidoc.ActiveView
    except Exception:
        active_view = None
    if active_view is None:
        return state

    try:
        detail_param = active_view.get_Parameter(BuiltInParameter.VIEW_DETAIL_LEVEL)
    except Exception:
        detail_param = None
    if detail_param is None or detail_param.IsReadOnly:
        return state

    try:
        current_value = detail_param.AsInteger()
        fine_value = int(ViewDetailLevel.Fine)
    except Exception:
        return state

    state['param'] = detail_param
    state['previous'] = current_value
    if current_value == fine_value:
        return state

    tx = Transaction(doc, 'Настройка детализации вида (временно)')
    try:
        tx.Start()
        detail_param.Set(fine_value)
        tx.Commit()
        state['changed'] = True
    except Exception:
        try:
            tx.RollBack()
        except Exception:
            pass
    return state


def _restore_view_detail(state):
    if not state or not state.get('changed'):
        return

    detail_param = state.get('param')
    previous_value = state.get('previous')
    if detail_param is None or previous_value is None:
        return

    tx = Transaction(doc, 'Возврат детализации вида')
    try:
        tx.Start()
        detail_param.Set(previous_value)
        tx.Commit()
    except Exception:
        try:
            tx.RollBack()
        except Exception:
            pass

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


def _get_material_name(material_id):
    if material_id is None or not isinstance(material_id, ElementId):
        return u''

    try:
        if material_id.IntegerValue < 1:
            return u''
    except Exception:
        pass

    try:
        material = doc.GetElement(material_id)
    except Exception:
        material = None

    name = _get_element_name(material)
    if name:
        return name

    try:
        return u'ID {}'.format(material_id.IntegerValue)
    except Exception:
        return u''



def _collect_target_walls():
    walls = []
    seen_ids = set()
    non_wall_ids = []
    auto_host_pairs = []

    def _register_wall(wall):
        if wall is None:
            return False
        try:
            wall_id = wall.Id.IntegerValue
        except Exception:
            return False
        if wall_id in seen_ids:
            return True
        walls.append(wall)
        seen_ids.add(wall_id)
        return True

    def _resolve_host_wall(element):
        if element is None:
            return None
        if isinstance(element, Wall):
            return element

        candidates = []
        host = getattr(element, 'Host', None)
        if host is not None:
            candidates.append(host)

        for attr in ('HostElementId', 'HostId'):
            host_id = getattr(element, attr, None)
            if host_id and isinstance(host_id, ElementId) and host_id.IntegerValue > 0:
                try:
                    host_element = doc.GetElement(host_id)
                except Exception:
                    host_element = None
                if host_element is not None:
                    candidates.append(host_element)

        if isinstance(element, FamilyInstance):
            try:
                super_component = element.SuperComponent
            except Exception:
                super_component = None
            if super_component is not None and super_component != element:
                candidates.append(super_component)

        for candidate in candidates:
            if candidate is None:
                continue
            if isinstance(candidate, ElementId):
                try:
                    candidate = doc.GetElement(candidate)
                except Exception:
                    candidate = None
            if isinstance(candidate, Wall):
                return candidate
        return None

    def _process_element(element, id_value):
        wall = _resolve_host_wall(element)
        if wall is not None:
            registered = _register_wall(wall)
            if registered and not isinstance(element, Wall) and id_value is not None:
                try:
                    auto_host_pairs.append((int(id_value), wall.Id.IntegerValue))
                except Exception:
                    auto_host_pairs.append((id_value, wall.Id.IntegerValue))
            return

        if id_value is None:
            try:
                id_value = element.Id.IntegerValue
            except Exception:
                id_value = None
        if id_value is not None:
            non_wall_ids.append(int(id_value))

    def _log_skipped(source):
        if not non_wall_ids:
            return
        descriptor = 'выбранный' if source == 'selected' else 'указанный'
        if len(non_wall_ids) == 1:
            logger.warning('Пропускаю %s элемент %s: это не стена.', descriptor, non_wall_ids[0])
        else:
            preview = ', '.join(str(item) for item in non_wall_ids[:5])
            if len(non_wall_ids) > 5:
                preview += ', ...'
            logger.warning('Пропускаю элементы, не являющиеся стенами (%s шт.): %s', len(non_wall_ids), preview)

    def _log_auto_hosts():
        if not auto_host_pairs:
            return
        preview = ', '.join('{}->{}'.format(src, dst) for src, dst in auto_host_pairs[:5])
        if len(auto_host_pairs) > 5:
            preview += ', ...'
        logger.debug('Автоматически добавлены стены-хосты для размещённых элементов: %s', preview)

    selected_ids = list(uidoc.Selection.GetElementIds())
    if selected_ids:
        for element_id in selected_ids:
            element = doc.GetElement(element_id)
            if element is None:
                continue
            try:
                id_value = element_id.IntegerValue
            except Exception:
                id_value = None
            _process_element(element, id_value)
        _log_skipped('selected')
        _log_auto_hosts()
        return walls

    try:
        references = uidoc.Selection.PickObjects(
            ObjectType.Element,
            'Выберите стены для разбиения на слои'
        )
    except Exception:
        return []

    if not references:
        return walls

    for reference in references:
        if reference is None:
            continue
        try:
            element = doc.GetElement(reference.ElementId)
        except Exception:
            element = None
        if element is None:
            continue
        try:
            id_value = reference.ElementId.IntegerValue
        except Exception:
            id_value = None
        _process_element(element, id_value)

    _log_skipped('picked')
    _log_auto_hosts()
    return walls


def _structure_layers_data(structure):
    layer_items = _layers_to_sequence(structure.GetLayers())
    layer_count = len(layer_items)
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

            if (
                candidate
                and isinstance(candidate, ElementId)
                and candidate.IntegerValue > 0
            ):
                material_id = candidate
                break

        material_name = _get_material_name(material_id)

        if (
            (material_id is None)
            or (not isinstance(material_id, ElementId))
            or material_id.IntegerValue < 1
        ):
            try:
                candidate = structure.GetMaterialId(idx)
            except Exception:
                candidate = None

            if (
                candidate
                and isinstance(candidate, ElementId)
                and candidate.IntegerValue > 0
            ):
                material_id = candidate
                material_name = _get_material_name(material_id)

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
            'material_name': material_name,
            'function': function,
            'is_core': is_core,
            'is_first': idx == 0,
            'is_last': idx == (layer_count - 1),
        })

    total_width = position
    core_start = data[first_core]['start'] if data and first_core not in (-1, None) else 0.0
    core_end = data[last_core]['end'] if data and last_core not in (-1, None) else total_width

    return data, total_width, core_start, core_end


def _select_host_layer(layer_data):
    choices = [_LayerChoice(info) for info in layer_data]
    if not choices:
        return None

    default_choice = None
    for choice in choices:
        if choice.layer_info.get('function') == MaterialFunctionAssignment.Structure:
            default_choice = choice
            break

    dialog = _LayerSelectionDialog(choices, default_choice)
    selection = dialog.show_dialog()

    if isinstance(selection, _LayerChoice):
        try:
            return dict(selection.layer_info)
        except Exception:
            return selection.layer_info

    return None


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
        inverted = vector.Negate()
        if inverted is not None:
            return inverted
    except Exception:
        pass

    try:
        x = -getattr(vector, 'X', 0.0)
        y = -getattr(vector, 'Y', 0.0)
        z = -getattr(vector, 'Z', 0.0)
        return XYZ(x, y, z)
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


def _shrink_curve(curve, shrink_distance, min_length=1e-6):
    """Возвращает копию кривой, укороченной на указанную величину."""

    if curve is None:
        return None

    try:
        shrink = abs(float(shrink_distance or 0.0))
    except Exception:
        shrink = 0.0

    if shrink <= min_length:
        return curve

    try:
        length = float(curve.Length)
    except Exception:
        try:
            length = float(curve.ApproximateLength)
        except Exception:
            length = None

    if length is None or length <= min_length:
        return curve

    max_shrink = max(0.0, length - min_length)
    if max_shrink <= min_length:
        return curve

    shrink = min(shrink, max_shrink)
    half_shrink = shrink / 2.0

    try:
        start_param = curve.GetEndParameter(0)
        end_param = curve.GetEndParameter(1)
        param_range = end_param - start_param
        if param_range > 0:
            param_delta = (half_shrink / length) * param_range
            new_start = start_param + param_delta
            new_end = end_param - param_delta
            if new_end > new_start:
                trimmed = curve.CreateTrimmedCurve(new_start, new_end)
                if trimmed:
                    return trimmed
    except Exception:
        pass

    try:
        start = curve.GetEndPoint(0)
        end = curve.GetEndPoint(1)
        direction = end - start
        if direction.IsZeroLength():
            return curve
        direction = direction.Normalize()
        new_start = start + direction.Multiply(half_shrink)
        new_end = end - direction.Multiply(half_shrink)
        return Line.CreateBound(new_start, new_end)
    except Exception:
        return curve


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

    base_name = _get_element_name(source_type)
    if not base_name:
        base_name = u"Слой стены"

    base_label = u"{} Слой {} {:.1f} мм".format(base_name, layer_info['index'], width_mm)
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


def _compute_wall_height_from_geometry(wall, curve=None):
    heights = []

    try:
        bbox = wall.get_BoundingBox(None)
    except Exception:
        bbox = None

    if bbox is not None:
        transform = getattr(bbox, 'Transform', None)
        min_pt = bbox.Min
        max_pt = bbox.Max
        if transform is not None:
            try:
                min_pt = transform.OfPoint(min_pt)
                max_pt = transform.OfPoint(max_pt)
            except Exception:
                pass
        if min_pt is not None and max_pt is not None:
            try:
                heights.append(abs(max_pt.Z - min_pt.Z))
            except Exception:
                pass

    if curve is None:
        location = getattr(wall, 'Location', None)
        if isinstance(location, LocationCurve):
            curve = location.Curve

    if curve is not None:
        try:
            start_z = curve.GetEndPoint(0).Z
            end_z = curve.GetEndPoint(1).Z
            heights.append(abs(end_z - start_z))
        except Exception:
            pass

    valid_heights = [h for h in heights if h is not None and h > 0]
    if valid_heights:
        return max(valid_heights)
    return None


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

    height = None
    height_param = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM)
    if height_param is not None:
        has_value = True
        try:
            has_value = height_param.HasValue
        except Exception:
            pass
        if has_value:
            try:
                candidate_height = height_param.AsDouble()
            except Exception:
                candidate_height = None
            if candidate_height is not None and candidate_height > 0:
                height = candidate_height

    if height is None or height <= 0:
        geometry_height = _compute_wall_height_from_geometry(wall, curve)
        if geometry_height is not None and geometry_height > 0:
            height = geometry_height

    if height is None or height <= 0:
        logger.warning(
            u'Высота стены %s недоступна или неположительна (%.4f). '
            u'Используем минимальное допустимое значение Revit.',
            wall.Id,
            height or 0.0,
        )
        height = _REVIT_MIN_DIMENSION
    elif height < _REVIT_MIN_DIMENSION:
        logger.warning(
            u'Высота стены %s (%.4f) меньше минимально допустимой. '
            u'Используем минимальное значение Revit.',
            wall.Id,
            height,
        )
        height = _REVIT_MIN_DIMENSION

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


def _extract_opening_points(instance, target_wall=None):
    if instance is None:
        return []

    try:
        opening_filter = ElementClassFilter(Opening)
        opening_ids = instance.GetDependentElements(opening_filter)
    except Exception:
        opening_ids = []

    if not opening_ids:
        return []

    target_id = _element_id_to_int(getattr(target_wall, 'Id', None)) if target_wall is not None else None

    points = []
    for open_id in opening_ids:
        opening = doc.GetElement(open_id)
        if opening is None:
            continue

        host = getattr(opening, 'Host', None)
        host_id = _element_id_to_int(getattr(host, 'Id', None)) if host is not None else _element_id_to_int(getattr(opening, 'HostId', None))
        if target_id is not None and host_id is not None and host_id != target_id:
            continue

        candidate_points = []
        curves = getattr(opening, 'BoundaryCurves', None)
        if curves:
            for curve in curves:
                for end_idx in (0, 1):
                    try:
                        candidate_points.append(curve.GetEndPoint(end_idx))
                    except Exception:
                        pass

        if not candidate_points:
            try:
                bbox = opening.get_BoundingBox(None)
            except Exception:
                bbox = None
            if bbox is not None:
                transform = bbox.Transform or Transform.Identity
                min_pt = bbox.Min
                max_pt = bbox.Max
                raw_corners = [
                    XYZ(min_pt.X, min_pt.Y, min_pt.Z),
                    XYZ(max_pt.X, min_pt.Y, min_pt.Z),
                    XYZ(min_pt.X, max_pt.Y, min_pt.Z),
                    XYZ(max_pt.X, max_pt.Y, min_pt.Z),
                    XYZ(min_pt.X, min_pt.Y, max_pt.Z),
                    XYZ(max_pt.X, min_pt.Y, max_pt.Z),
                    XYZ(min_pt.X, max_pt.Y, max_pt.Z),
                    XYZ(max_pt.X, max_pt.Y, max_pt.Z),
                ]
                candidate_points = [transform.OfPoint(pt) for pt in raw_corners]

        if candidate_points:
            points = candidate_points
            break

    return points


def _measure_opening_width(points, wall):
    if not points or wall is None:
        return None

    wall_location = getattr(wall, 'Location', None)
    if not isinstance(wall_location, LocationCurve):
        return None

    wall_curve = wall_location.Curve
    if wall_curve is None:
        return None

    try:
        direction_vec = wall_curve.GetEndPoint(1) - wall_curve.GetEndPoint(0)
        if direction_vec.IsZeroLength():
            return None
        direction = direction_vec.Normalize()
    except Exception:
        return None

    projections = [direction.DotProduct(pt) for pt in points]
    if not projections:
        return None

    return max(projections) - min(projections)


def _ensure_opening_for_instance(
        instance,
        wall,
        margin=_OPENING_MARGIN,
        reference_points=None,
        preferred_width=None,
        geometry_cache=None,
):
    if instance is None or wall is None:
        return False

    wall_location = getattr(wall, 'Location', None)
    if not isinstance(wall_location, LocationCurve):
        return False

    wall_curve = wall_location.Curve
    if wall_curve is None:
        return False

    try:
        direction_vec = wall_curve.GetEndPoint(1) - wall_curve.GetEndPoint(0)
        if direction_vec.IsZeroLength():
            return False
        wall_direction = direction_vec.Normalize()
    except Exception:
        return False

    normal = getattr(wall, 'Orientation', None)
    if normal is None or normal.IsZeroLength():
        normal = getattr(instance, 'FacingOrientation', None)
        if normal is None or normal.IsZeroLength():
            return False
        normal = normal.Normalize()
    else:
        normal = normal.Normalize()

    points = []
    if reference_points:
        points = list(reference_points)
    if not points:
        points = _extract_opening_points(instance, wall)

    if not points:
        bbox = None
        try:
            bbox = instance.get_BoundingBox(None)
        except Exception:
            bbox = None
        if bbox is not None:
            transform = bbox.Transform or Transform.Identity
            min_pt = bbox.Min
            max_pt = bbox.Max
            raw_corners = [
                XYZ(min_pt.X, min_pt.Y, min_pt.Z),
                XYZ(max_pt.X, min_pt.Y, min_pt.Z),
                XYZ(min_pt.X, max_pt.Y, min_pt.Z),
                XYZ(max_pt.X, max_pt.Y, min_pt.Z),
                XYZ(min_pt.X, min_pt.Y, max_pt.Z),
                XYZ(max_pt.X, min_pt.Y, max_pt.Z),
                XYZ(min_pt.X, max_pt.Y, max_pt.Z),
                XYZ(max_pt.X, max_pt.Y, max_pt.Z),
            ]
            points = [transform.OfPoint(pt) for pt in raw_corners]

    if not points:
        return False

    projections = [wall_direction.DotProduct(pt) for pt in points]
    min_proj_points = min(projections)
    max_proj_points = max(projections)
    width_from_points = max_proj_points - min_proj_points

    heights = [pt.Z for pt in points]
    bottom_z_points = min(heights)
    top_z_points = max(heights)
    height_from_points = top_z_points - bottom_z_points

    geometry_info = geometry_cache
    if geometry_info is None:
        geometry_info = _get_instance_geometry_metrics(instance, wall_direction)

    geometry_width = geometry_info.get('width') if geometry_info else None
    geometry_bottom = geometry_info.get('bottom') if geometry_info else None
    geometry_top = geometry_info.get('top') if geometry_info else None
    geometry_height = geometry_info.get('height') if geometry_info else None

    host_reference_width = None
    if reference_points:
        host_reference_width = _measure_opening_width(reference_points, wall)

    door_height_param = _get_param_double(instance, BuiltInParameter.DOOR_HEIGHT)
    sill_height = _get_param_double(instance, BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM)
    head_height = _get_param_double(instance, BuiltInParameter.INSTANCE_HEAD_HEIGHT_PARAM)

    level_elevation = None
    try:
        level = doc.GetElement(instance.LevelId)
        if level is not None:
            level_elevation = getattr(level, 'Elevation', None)
    except Exception:
        level_elevation = None

    bottom_from_params = None
    if level_elevation is not None and sill_height is not None:
        bottom_from_params = level_elevation + sill_height

    top_from_params = None
    if level_elevation is not None and head_height is not None:
        top_from_params = level_elevation + head_height
    elif bottom_from_params is not None and door_height_param is not None:
        top_from_params = bottom_from_params + door_height_param

    if bottom_from_params is None:
        bottom_from_params = bottom_z_points
    if top_from_params is None:
        top_from_params = top_z_points

    bottom_z = bottom_from_params
    top_z = top_from_params
    if geometry_height is not None and geometry_height > _WIDTH_EPS:
        if geometry_bottom is not None:
            bottom_z = geometry_bottom
        if geometry_top is not None:
            top_z = geometry_top

    effective_preferred = preferred_width
    if effective_preferred is None or effective_preferred <= _WIDTH_EPS:
        fallback_width = _get_instance_clear_width(instance)
        if fallback_width is not None and fallback_width > _WIDTH_EPS:
            effective_preferred = fallback_width

    width = None
    if effective_preferred is not None and effective_preferred > _WIDTH_EPS:
        width = effective_preferred
    else:
        width_candidates = [
            geometry_width,
            host_reference_width,
            width_from_points,
        ]
        width_values = [val for val in width_candidates if val is not None and val > _WIDTH_EPS]
        if not width_values:
            return False
        width = min(width_values)

    height_candidates = [
        geometry_height,
        (top_z - bottom_z) if top_z is not None and bottom_z is not None else None,
        height_from_points,
    ]
    height_values = [val for val in height_candidates if val is not None and val > _WIDTH_EPS]
    if not height_values:
        return False
    height = max(height_values)

    width = width + 2.0 * margin
    height = height + 2.0 * margin

    location = getattr(instance, 'Location', None)
    reference_point = None
    if isinstance(location, LocationPoint):
        reference_point = location.Point
    elif isinstance(location, LocationCurve):
        try:
            reference_point = location.Curve.Evaluate(0.5, True)
        except Exception:
            reference_point = None

    bbox = None
    if reference_point is None:
        try:
            bbox = instance.get_BoundingBox(None)
        except Exception:
            bbox = None
        if bbox is not None:
            transform = bbox.Transform or Transform.Identity
            center_local = XYZ(
                (bbox.Min.X + bbox.Max.X) / 2.0,
                (bbox.Min.Y + bbox.Max.Y) / 2.0,
                (bbox.Min.Z + bbox.Max.Z) / 2.0,
            )
            reference_point = transform.OfPoint(center_local)
    if reference_point is None:
        reference_point = XYZ(
            sum(pt.X for pt in points) / len(points),
            sum(pt.Y for pt in points) / len(points),
            sum(pt.Z for pt in points) / len(points),
        )

    plane_origin = wall_curve.GetEndPoint(0)
    vector_to_plane = reference_point - plane_origin
    distance = normal.DotProduct(vector_to_plane)
    center_on_plane = reference_point - normal.Multiply(distance)

    projection = wall_curve.Project(center_on_plane)
    if projection is not None:
        try:
            center_along_curve = projection.XYZPoint
        except Exception:
            center_along_curve = center_on_plane
    else:
        center_along_curve = center_on_plane

    center_height = (bottom_z + top_z) / 2.0 if bottom_z is not None and top_z is not None else reference_point.Z
    center_point = XYZ(center_along_curve.X, center_along_curve.Y, center_height)

    half_width_vec = wall_direction.Multiply(width / 2.0)
    half_height_vec = XYZ.BasisZ.Multiply(height / 2.0)

    bottom_left = center_point - half_width_vec - half_height_vec
    top_right = center_point + half_width_vec + half_height_vec

    try:
        doc.Create.NewOpening(wall, bottom_left, top_right)
        return True
    except Exception as exc:
        logger.debug('Не удалось создать дополнительный проём в стене %s: %s', wall.Id.IntegerValue, exc)
        return False


def _rehost_instances(instances, new_host_wall, other_walls=None):
    if new_host_wall is None:
        return

    if other_walls is None:
        other_walls = []

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
            if info.get('hand_flipped') and hasattr(new_inst, 'HandFlipped') and not new_inst.HandFlipped:
                new_inst.flipHand()
        except Exception:
            pass

        try:
            if info.get('face_flipped') and hasattr(new_inst, 'FacingFlipped') and not new_inst.FacingFlipped:
                new_inst.flipFacing()
        except Exception:
            pass

        try:
            doc.Regenerate()
        except Exception:
            pass

        host_direction = None
        host_location = getattr(new_host_wall, 'Location', None)
        if isinstance(host_location, LocationCurve):
            host_curve = host_location.Curve
            if host_curve is not None:
                try:
                    host_vec = host_curve.GetEndPoint(1) - host_curve.GetEndPoint(0)
                    if not host_vec.IsZeroLength():
                        host_direction = host_vec.Normalize()
                except Exception:
                    host_direction = None

        geometry_metrics = _get_instance_geometry_metrics(new_inst, host_direction)

        host_opening_points = _extract_opening_points(new_inst, new_host_wall)

        width_from_params = _get_instance_clear_width(new_inst)
        width_from_geometry = geometry_metrics.get('width') if geometry_metrics else None

        preferred_width = None
        if width_from_params is not None and width_from_params > _WIDTH_EPS:
            preferred_width = width_from_params
        elif width_from_geometry is not None and width_from_geometry > _WIDTH_EPS:
            preferred_width = width_from_geometry


        fallback_walls = []
        for extra_wall in other_walls or []:
            if extra_wall is None or extra_wall.Id == new_host_wall.Id:
                continue

            void_cut_done = False
            if _CAN_ADD_VOID_CUT and _ADD_INSTANCE_VOID_CUT:
                try:
                    if _CAN_ADD_VOID_CUT(doc, new_inst, extra_wall):
                        _ADD_INSTANCE_VOID_CUT(doc, new_inst, extra_wall)
                        void_cut_done = True
                except Exception:
                    try:
                        extra_id = extra_wall.Id.IntegerValue if extra_wall else None
                    except Exception:
                        extra_id = None
                    logger.debug(
                        'Не удалось прорезать стену %s пустотой экземпляра %s',
                        extra_id,
                        getattr(getattr(new_inst, 'Id', None), 'IntegerValue', None),
                    )

            if void_cut_done:
                continue

            fallback_walls.append(extra_wall)

        for extra_wall in fallback_walls:
            try:
                created = _ensure_opening_for_instance(
                    new_inst,
                    extra_wall,
                    reference_points=host_opening_points,
                    preferred_width=preferred_width,
                    geometry_cache=geometry_metrics,
                )
                if not created:
                    logger.debug(
                        'Не удалось создать проём в стене %s через прямоугольное отверстие',
                        getattr(getattr(extra_wall, 'Id', None), 'IntegerValue', None),
                    )
            except Exception as exc:
                logger.debug(
                    'Ошибка при создании проёма в стене %s: %s',
                    getattr(getattr(extra_wall, 'Id', None), 'IntegerValue', None),
                    exc,
                )

        try:
            doc.Delete(info['id'])
        except Exception:
            pass


def _breakup_wall(wall, show_alert=True):
    wall_id = None
    try:
        wall_id = wall.Id.IntegerValue
    except Exception:
        pass

    wall_type_id = None
    try:
        wall_type_id = _element_id_to_int(wall.WallType.Id)
    except Exception:
        wall_type_id = None

    joined_wall_ids = []

    def _handle_failure(message, level='warning', status='error'):
        if show_alert:
            forms.alert(message)
        else:
            if level == 'error':
                logger.error(message)
            elif level == 'info':
                logger.info(message)
            else:
                logger.warning(message)
        return {
            'status': status,
            'created': 0,
            'message': message,
            'wall_id': wall_id,
        }

    structure, layers = _collect_structure(wall)
    if not layers:
        return _handle_failure('Для стены {} не найдены слои составной конструкции.'.format(wall_id))

    layer_data, total_width, core_start, core_end = _structure_layers_data(structure)
    layer_data = [item for item in layer_data if item['width'] > _WIDTH_EPS]
    if layer_data:
        last_index = len(layer_data) - 1
        for idx, item in enumerate(layer_data):
            item['is_first'] = idx == 0
            item['is_last'] = idx == last_index
    if not layer_data:
        return _handle_failure('Для стены {} нет слоёв с ненулевой толщиной.'.format(wall_id))

    joined_wall_ids = _collect_joined_wall_ids(wall, wall_type_id, layer_data)

    hosted_instances = _collect_hosted_instances(wall)
    host_layer_info = None
    selected_layer_index = None
    preview_state = None
    if hosted_instances:
        try:
            uidoc.ShowElements(wall.Id)
        except Exception:
            pass

        try:
            selection_ids = List[ElementId]()
            selection_ids.Add(wall.Id)
            uidoc.Selection.SetElementIds(selection_ids)
        except Exception:
            pass
        try:
            uidoc.RefreshActiveView()
        except Exception:
            pass

        preview_state = _adjust_view_detail_for_preview()
        try:
            host_layer_info = _select_host_layer(layer_data)
        finally:
            _restore_view_detail(preview_state)
        if not host_layer_info:
            return _handle_failure('Операция отменена пользователем.', level='info', status='cancelled')

        selected_layer_index = host_layer_info.get('index')

    try:
        context = _build_wall_context(wall, structure, total_width, core_start, core_end)
    except ValueError as exc:
        return _handle_failure(str(exc))

    orientation = _ensure_orientation_vector(context, context['curve'])
    inward = _negate_vector(orientation)
    try:
        inward = inward.Normalize()
    except Exception:
        pass

    base_curve = context['curve']
    base_level_id = context['base_level_id']

    created_walls = []
    produced_layers = []
    t = Transaction(doc, 'Разделение стены на слои')
    t.Start()
    try:
        _ensure_join_controls(wall)
        _suppress_auto_join(wall)

        for layer_info in layer_data:
            layer_type = _clone_wall_type_for_layer(wall.WallType, layer_info)
            if layer_type is None:
                continue

            layer_center = (layer_info['start'] + layer_info['end']) / 2.0
            offset_center = layer_center - context['reference_offset']
            layer_info['offset'] = offset_center
            layer_info['reference_offset'] = context['reference_offset']
            translation_vector = _scale_vector(inward, offset_center)
            layer_width = layer_info.get('width')
            if layer_info.get('is_first') or layer_info.get('is_last'):
                shrink_distance = 0.0
            else:
                shrink_distance = layer_width
            adjusted_curve = _shrink_curve(base_curve, shrink_distance) or base_curve
            placement_curve = adjusted_curve.CreateTransformed(
                Transform.CreateTranslation(translation_vector)
            )

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

            existing_parts = list(created_walls)
            created_walls.append(new_wall)
            produced_layers.append({
                'wall': new_wall,
                'info': dict(layer_info),
                'offset': offset_center,
                'reference_offset': context['reference_offset'],
            })
            logger.debug('Создана стена %s для слоя %s', new_wall.Id.IntegerValue, layer_info['index'])

            _prepare_wall_for_manual_join(new_wall, existing_parts, wall)

        if not created_walls:
            t.RollBack()
            return _handle_failure('Для стены {} не удалось создать ни одной разделённой стены.'.format(wall_id))

        host_wall = None
        for record in produced_layers:
            layer_info = record.get('info') or {}
            wall_part = record.get('wall')
            if wall_part is None:
                continue
            if selected_layer_index is not None and layer_info.get('index') == selected_layer_index:
                host_wall = wall_part
                break

        if host_wall is None:
            for record in produced_layers:
                layer_info = record.get('info') or {}
                wall_part = record.get('wall')
                if wall_part is None:
                    continue
                if layer_info.get('function') == MaterialFunctionAssignment.Structure:
                    host_wall = wall_part
                    break

        if host_wall is None:
            host_wall = produced_layers[0]['wall']

        host_index = None
        for idx, record in enumerate(produced_layers):
            wall_part = record.get('wall')
            if wall_part is not None and wall_part.Id == host_wall.Id:
                host_index = idx
                break

        if host_index is None:
            host_index = 0

        preceding_walls = []
        for idx, record in enumerate(produced_layers):
            wall_part = record.get('wall')
            if wall_part is None or wall_part.Id == host_wall.Id:
                continue
            if idx < host_index:
                preceding_walls.append(wall_part)

        _rehost_instances(hosted_instances, host_wall, preceding_walls)

        for i in range(len(created_walls)):
            for j in range(i + 1, len(created_walls)):
                first_id = getattr(created_walls[i], 'Id', None)
                second_id = getattr(created_walls[j], 'Id', None)
                if first_id is None or second_id is None:
                    continue
                _join_two_walls(first_id, second_id)

        try:
            doc.Regenerate()
        except Exception:
            pass

        try:
            _handle_layer_joins(
                wall_id,
                wall_type_id,
                produced_layers,
                joined_wall_ids,
            )
        except Exception as exc:
            logger.debug('Не удалось обработать соединения слоев для стены %s: %s', wall_id, exc)

        try:
            doc.Delete(wall.Id)
        except Exception as exc:
            logger.warning('Не удалось удалить исходную стену %s: %s', wall_id, exc)
            t.RollBack()
            return _handle_failure('Не удалось удалить исходную стену {}.'.format(wall_id))

        t.Commit()
    except Exception as exc:
        logger.exception('Непредвиденная ошибка при разбиении стены %s: %s', wall_id, exc)
        t.RollBack()
        return _handle_failure('Непредвиденная ошибка при обработке стены {}.'.format(wall_id), level='error')

    success_message = 'Создано {} стен-слоёв из стены {}.'.format(len(created_walls), wall_id)
    if show_alert:
        forms.alert(success_message)
    else:
        logger.info(success_message)

    return {
        'status': 'success',
        'created': len(created_walls),
        'message': success_message,
        'wall_id': wall_id,
    }


def main():
    _LAYER_JOIN_CACHE.clear()
    _PENDING_LAYER_JOINS.clear()

    walls = _collect_target_walls()
    if not walls:
        forms.alert('Не выбрано ни одной стены.')
        return

    multi_mode = len(walls) > 1
    results = []
    for wall in walls:
        result = _breakup_wall(wall, show_alert=not multi_mode)
        results.append(result)
        if result.get('status') == 'cancelled':
            break

    if multi_mode:
        success_results = [item for item in results if item.get('status') == 'success']
        failure_results = [item for item in results if item.get('status') != 'success']
        total_created = sum(item.get('created', 0) for item in success_results)

        summary_lines = [
            'Обработано стен: {}'.format(len(results)),
            'Стены успешно разделены: {}'.format(len(success_results)),
            'Создано новых стен-слоёв: {}'.format(total_created),
        ]
        if failure_results:
            summary_lines.append('Пропущено или с ошибкой: {}'.format(len(failure_results)))
        forms.alert('\n'.join(summary_lines))


if __name__ == '__main__':
    main()

