# -*- coding: utf-8 -*-
"""Разделение составной стены на отдельные однослойные стены.

Инструмент для pyRevit. Алгоритм работы:
1. Пользователь выбирает составную стену.
2. Скрипт анализирует слои выбранного типа стены.
3. Пользователь выбирает какие слои отделить и к каким типам стен их привязать.
4. Для выбранных слоёв создаются новые типы стен (либо используются существующие).
5. Создаются новые стены с корректным смещением относительно исходной стены.
6. Переносятся окна/двери на выбранную новую стену.
7. Параметры исходной стены копируются в новые.
8. Исходная стена удаляется, а новые стены объединяются Join Geometry.
"""

import math

from pyrevit import forms, revit, script
from pyrevit.forms import TemplateListItem

from Autodesk.Revit import DB

__title__ = "Разделить стену по слоям"
__author__ = "pw-team"


class LayerInfo(object):
    """Информация об одном слое составной стены."""

    def __init__(
        self,
        index,
        name,
        width,
        material_id,
        function,
        offset,
        is_core,
    ):
        self.index = index
        self.name = name
        self.width = width
        self.material_id = material_id
        self.function = function
        self.offset = offset
        self.is_core = is_core

        self.wall_type = None
        self.new_wall = None
        self.translation = None

    @property
    def width_mm(self):
        return DB.UnitUtils.ConvertFromInternalUnits(
            self.width, DB.UnitTypeId.Millimeters
        )

    @property
    def display_name(self):
        return "{} ({:.1f} мм)".format(self.name, self.width_mm)


def _pick_wall():
    class WallSelectionFilter(DB.ISelectionFilter):
        def AllowElement(self, element):
            return isinstance(element, DB.Wall)

        def AllowReference(self, reference, position):
            return False

    message = "Выберите составную стену"
    ref = revit.pick_element(message=message, selection_filter=WallSelectionFilter())
    if not ref:
        forms.alert("Стена не выбрана.", exitscript=True)
    wall = revit.doc.GetElement(ref.ElementId)
    if not isinstance(wall, DB.Wall):
        forms.alert("Выбранный элемент не является стеной.", exitscript=True)
    return wall


def _get_layers(wall):
    wall_type = revit.doc.GetElement(wall.GetTypeId())
    if not isinstance(wall_type, DB.WallType) or wall_type.Kind != DB.WallKind.Basic:
        forms.alert(
            "Выбрана несоставная стена. Инструмент работает только с обычными (Basic) стенами.",
            exitscript=True,
        )

    structure = wall_type.GetCompoundStructure()
    if not structure:
        forms.alert("Тип стены не содержит слоёв.", exitscript=True)

    layers_info = []
    layer_count = structure.LayerCount
    for index in range(layer_count):
        width = structure.GetLayerWidth(index)
        if width <= 0:
            continue
        material_id = structure.GetMaterialId(index)
        material = revit.doc.GetElement(material_id) if material_id else None
        mat_name = material.Name if material else "Без материала"
        try:
            layer_name = structure.GetLayerName(index)
        except Exception:
            layer_name = mat_name
        if not layer_name:
            layer_name = mat_name
        function = structure.GetLayerFunction(index)
        offset = structure.GetLayerOffset(index)
        is_core = structure.IsCoreLayer(index)
        layers_info.append(
            LayerInfo(
                index=index,
                name=layer_name,
                width=width,
                material_id=material_id,
                function=function,
                offset=offset,
                is_core=is_core,
            )
        )

    if not layers_info:
        forms.alert("Не найдено слоёв подходящих для разделения.", exitscript=True)

    return layers_info


def _select_layers(layers):
    items = []
    for layer in layers:
        label = layer.display_name
        if layer.is_core:
            label += " [CORE]"
        items.append(TemplateListItem(label, layer))

    selected = forms.SelectFromList.show(
        items,
        title="Выберите слои для разделения",
        multiselect=True,
        button_name="Продолжить",
    )
    if not selected:
        forms.alert("Слои не выбраны.", exitscript=True)

    return [item.value for item in selected]


def _collect_wall_types():
    collector = DB.FilteredElementCollector(revit.doc).OfClass(DB.WallType)
    result = {}
    for wall_type in collector:
        if wall_type.Kind != DB.WallKind.Basic:
            continue
        structure = wall_type.GetCompoundStructure()
        if not structure:
            continue
        result[wall_type.Id.IntegerValue] = wall_type
    return result


def _wall_type_thickness(wall_type):
    structure = wall_type.GetCompoundStructure()
    if not structure:
        return 0.0
    total = 0.0
    for index in range(structure.LayerCount):
        total += structure.GetLayerWidth(index)
    return total


def _ask_mapping(layers, wall_types):
    tolerance = DB.UnitUtils.ConvertToInternalUnits(1.0, DB.UnitTypeId.Millimeters)

    components = []
    mapping_keys = []
    for layer in layers:
        options = [TemplateListItem("Создать новый тип", None)]
        for wt in wall_types.values():
            thickness = _wall_type_thickness(wt)
            if math.fabs(thickness - layer.width) <= tolerance:
                item_text = "{} ({:.1f} мм)".format(
                    wt.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString(),
                    DB.UnitUtils.ConvertFromInternalUnits(thickness, DB.UnitTypeId.Millimeters),
                )
                options.append(TemplateListItem(item_text, wt.Id))
        label = forms.Label("{}".format(layer.display_name))
        combo_name = "layer_{}".format(layer.index)
        combo = forms.ComboBox(combo_name, options, default=options[0])
        components.append(label)
        components.append(combo)
        mapping_keys.append(combo_name)

    host_options = [TemplateListItem(layer.display_name, layer.index) for layer in layers]
    components.append(forms.Label("Слой для переноса окон и дверей"))
    components.append(forms.ComboBox("host_layer", host_options, default=host_options[0]))
    components.append(forms.Button("Готово"))

    form = forms.FlexForm("Замена типов стен", components)
    form.show()
    if not form.values:
        forms.alert("Действие отменено пользователем.", exitscript=True)

    for key, layer in zip(mapping_keys, layers):
        selection = form.values.get(key)
        if isinstance(selection, TemplateListItem):
            selection_value = selection.value
        else:
            selection_value = selection
        if isinstance(selection_value, DB.ElementId):
            layer.wall_type = revit.doc.GetElement(selection_value)
        elif selection_value is None:
            layer.wall_type = None
        else:
            layer.wall_type = revit.doc.GetElement(DB.ElementId(int(selection_value)))

    host_selection = form.values.get("host_layer")
    if isinstance(host_selection, TemplateListItem):
        host_index = host_selection.value
    else:
        host_index = host_selection
    if host_index is None and layers:
        host_index = layers[0].index

    join_new = True

    return host_index, join_new


def _create_single_layer_type(base_type, layer):
    new_name = "{} | {}".format(base_type.Name, layer.name)

    collector = DB.FilteredElementCollector(revit.doc).OfClass(DB.WallType)
    for wt in collector:
        if wt.Name == new_name:
            return wt

    new_type = base_type.Duplicate(new_name)
    material_id = layer.material_id if layer.material_id else DB.ElementId.InvalidElementId
    structure = DB.CompoundStructure.CreateSimpleCompoundStructure(layer.width, material_id)
    structure.SetLayerFunction(0, layer.function)
    structure.SetNumberOfShellLayers(DB.ShellLayerType.Exterior, 0)
    structure.SetNumberOfShellLayers(DB.ShellLayerType.Interior, 0)
    new_type.SetCompoundStructure(structure)
    return new_type


def _copy_instance_parameters(source, target):
    source_params = {}
    for src_param in source.Parameters:
        source_params[src_param.Definition.Name] = src_param
    for param in target.Parameters:
        if param.IsReadOnly:
            continue
        src_param = source_params.get(param.Definition.Name)
        if not src_param:
            continue
        if src_param.StorageType != param.StorageType:
            continue
        try:
            if param.StorageType == DB.StorageType.Double:
                param.Set(src_param.AsDouble())
            elif param.StorageType == DB.StorageType.Integer:
                param.Set(src_param.AsInteger())
            elif param.StorageType == DB.StorageType.String:
                value = src_param.AsString()
                if value is not None:
                    param.Set(value)
            elif param.StorageType == DB.StorageType.ElementId:
                param.Set(src_param.AsElementId())
        except Exception:
            continue


def _create_new_wall(layer, base_wall, base_type):
    location = base_wall.Location
    if not isinstance(location, DB.LocationCurve):
        raise ValueError("Стену с криволинейной геометрией обработать не удалось.")
    curve = location.Curve

    if layer.wall_type is None:
        wall_type = _create_single_layer_type(base_type, layer)
        layer.wall_type = wall_type
    else:
        wall_type = layer.wall_type

    level_id = base_wall.LevelId
    flip = base_wall.Flipped

    new_wall = DB.Wall.Create(revit.doc, curve, wall_type.Id, level_id, 0.0, 0.0, flip, False)
    _copy_instance_parameters(base_wall, new_wall)
    new_wall.StructuralUsage = base_wall.StructuralUsage

    orientation = base_wall.Orientation
    translation = orientation.Multiply(layer.offset)
    if translation.GetLength() > 1e-9:
        DB.ElementTransformUtils.MoveElement(revit.doc, new_wall.Id, translation)

    layer.translation = translation
    layer.new_wall = new_wall
    return new_wall


def _collect_hosted_instances(wall):
    collector = DB.FilteredElementCollector(revit.doc).OfClass(DB.FamilyInstance)
    hosted = []
    for inst in collector:
        host = inst.Host
        if host and host.Id == wall.Id:
            hosted.append(inst)
    return hosted


def _create_hosted_instance(source_instance, host_wall, translation_vector):
    location = source_instance.Location
    if not isinstance(location, DB.LocationPoint):
        return None
    point = location.Point + translation_vector
    symbol = source_instance.Symbol
    level_id = source_instance.LevelId

    if not symbol.IsActive:
        symbol.Activate()
        revit.doc.Regenerate()

    new_instance = revit.doc.Create.NewFamilyInstance(
        point,
        symbol,
        host_wall,
        revit.doc.GetElement(level_id),
        DB.Structure.StructuralType.NonStructural,
    )

    new_location = new_instance.Location
    if isinstance(new_location, DB.LocationPoint):
        original_rotation = location.Rotation
        rotation_delta = original_rotation - new_location.Rotation
        if abs(rotation_delta) > 1e-6:
            axis = DB.Line.CreateBound(point, point + DB.XYZ.BasisZ)
            DB.ElementTransformUtils.RotateElement(revit.doc, new_instance.Id, axis, rotation_delta)

    if new_instance.HandFlipped != source_instance.HandFlipped:
        try:
            new_instance.flipHand()
        except Exception:
            pass
    if new_instance.FacingFlipped != source_instance.FacingFlipped:
        try:
            new_instance.flipFacing()
        except Exception:
            pass

    _copy_instance_parameters(source_instance, new_instance)
    return new_instance


def _join_walls(walls):
    count = len(walls)
    for i in range(count):
        for j in range(i + 1, count):
            first = walls[i]
            second = walls[j]
            try:
                DB.JoinGeometryUtils.JoinGeometry(revit.doc, first, second)
            except Exception:
                continue


def main():
    wall = _pick_wall()
    base_type = revit.doc.GetElement(wall.GetTypeId())
    layers = _get_layers(wall)
    selected_layers = _select_layers(layers)
    wall_types = _collect_wall_types()
    host_index, join_new = _ask_mapping(selected_layers, wall_types)

    hosted_instances = _collect_hosted_instances(wall)

    with revit.Transaction("Разделить стену по слоям"):
        new_walls = []
        for layer in selected_layers:
            new_wall = _create_new_wall(layer, wall, base_type)
            new_walls.append(new_wall)

        host_layer = next((l for l in selected_layers if l.index == host_index), None)
        if host_layer is None and selected_layers:
            host_layer = selected_layers[0]

        translation_vector = (
            host_layer.translation if host_layer and host_layer.translation else DB.XYZ.Zero
        )

        for inst in hosted_instances:
            _create_hosted_instance(inst, host_layer.new_wall if host_layer else new_walls[0], translation_vector)

        for inst in hosted_instances:
            try:
                revit.doc.Delete(inst.Id)
            except Exception:
                pass

        try:
            revit.doc.Delete(wall.Id)
        except Exception as err:
            forms.alert("Не удалось удалить исходную стену: {}".format(err))

        if join_new and len(new_walls) > 1:
            _join_walls(new_walls)

    script.get_output().print_md("**Готово.** Создано {} новых стен.".format(len(selected_layers)))


if __name__ == "__main__":
    main()
