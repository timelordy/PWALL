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

"""

import math

from pyrevit import forms, revit, script
from pyrevit.forms import TemplateListItem

from Autodesk.Revit import DB

__title__ = "Разделить стену по слоям"
__author__ = "pw-team"

# Универсальная обработка единиц измерения для разных версий Revit API
try:
    _MILLIMETER_UNIT = DB.UnitTypeId.Millimeters


    def _to_internal_mm(value):
        return DB.UnitUtils.ConvertToInternalUnits(value, _MILLIMETER_UNIT)


    def _from_internal_mm(value):
        return DB.UnitUtils.ConvertFromInternalUnits(value, _MILLIMETER_UNIT)

except AttributeError:
    _MILLIMETER_UNIT = DB.DisplayUnitType.DUT_MILLIMETERS


    def _to_internal_mm(value):
        return DB.UnitUtils.ConvertToInternalUnits(value, _MILLIMETER_UNIT)


    def _from_internal_mm(value):
        return DB.UnitUtils.ConvertFromInternalUnits(value, _MILLIMETER_UNIT)


class LayerInfo(object):
    """Информация об одном слое составной стены."""

    def __init__(
            self,
            index,
            name,
            width,
            material_id,
            function,
            offset_from_exterior,  # Изменено: смещение от внешней грани
            is_core,
    ):
        self.index = index
        self.name = name
        self.width = width
        self.material_id = material_id
        self.function = function
        self.offset_from_exterior = offset_from_exterior  # Изменено
        self.is_core = is_core

        self.wall_type = None
        self.new_wall = None
        self.translation = None

    @property
    def width_mm(self):
        return _from_internal_mm(self.width)

    @property
    def display_name(self):
        return "{} ({:.1f} мм)".format(self.name, self.width_mm)


def _pick_wall():
    """Выбор стены пользователем."""

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
    """Получение информации о слоях стены с правильным расчетом смещений."""
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

    # Расчет правильных смещений от внешней грани
    total_width = structure.GetWidth()
    current_offset = -total_width / 2.0  # Начинаем от внешней грани

    for index in range(layer_count):
        width = structure.GetLayerWidth(index)
        if width <= 0:
            continue

        # Смещение до центра текущего слоя
        offset_to_center = current_offset + width / 2.0

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
        is_core = structure.IsCoreLayer(index)

        layers_info.append(
            LayerInfo(
                index=index,
                name=layer_name,
                width=width,
                material_id=material_id,
                function=function,
                offset_from_exterior=offset_to_center,  # Используем правильное смещение
                is_core=is_core,
            )
        )

        current_offset += width  # Переходим к следующему слою

    if not layers_info:
        forms.alert("Не найдено слоёв подходящих для разделения.", exitscript=True)

    return layers_info


def _select_layers(layers):
    """Выбор слоёв для разделения."""
    items = []
    for layer in layers:
        label = layer.display_name
        if layer.is_core:
            label += " [НЕСУЩИЙ]"
        items.append(TemplateListItem(label, layer))

    selected = forms.SelectFromList.show(
        items,
        title="Выберите слои для разделения",
        multiselect=True,
        button_name="Продолжить",
    )
    if not selected:
        forms.alert("Слои не выбраны.", exitscript=True)

    return [item.value if hasattr(item, 'value') else item for item in selected]


def _collect_wall_types():
    """Сбор всех типов стен в проекте."""
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
    """Получение толщины типа стены."""
    structure = wall_type.GetCompoundStructure()
    if not structure:
        return 0.0
    return structure.GetWidth()


def _ask_mapping(layers, wall_types):
    """Интерфейс для выбора типов стен и настроек."""
    tolerance = _to_internal_mm(1.0)

    components = []
    mapping_keys = []

    for layer in layers:
        options = [TemplateListItem("Создать новый тип", None)]
        for wt in wall_types.values():
            thickness = _wall_type_thickness(wt)
            if math.fabs(thickness - layer.width) <= tolerance:
                item_text = "{} ({:.1f} мм)".format(
                    wt.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString(),
                    _from_internal_mm(thickness),
                )
                options.append(TemplateListItem(item_text, wt.Id))

        label = forms.Label("{}".format(layer.display_name))
        combo_name = "layer_{}".format(layer.index)
        combo = forms.ComboBox(combo_name, options, default=options[0])
        components.append(label)
        components.append(combo)
        mapping_keys.append(combo_name)

    host_options = [TemplateListItem(layer.display_name, layer.index) for layer in layers]
    components.append(forms.Label("Слой для переноса окон и дверей:"))
    components.append(forms.ComboBox("host_layer", host_options, default=host_options[0]))

    # Добавляем опцию соединения стен
    components.append(forms.CheckBox("join_walls", "Соединить новые стены между собой", default=True))
    components.append(forms.Button("Готово"))

    form = forms.FlexForm("Настройка разделения стен", components)
    form.show()

    if not form.values:
        forms.alert("Действие отменено пользователем.", exitscript=True)

    # Обработка выбранных типов стен
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

    # Обработка выбора слоя для окон/дверей
    host_selection = form.values.get("host_layer")
    if isinstance(host_selection, TemplateListItem):
        host_index = host_selection.value
    else:
        host_index = host_selection

    if host_index is None and layers:
        host_index = layers[0].index

    join_new = form.values.get("join_walls", True)

    return host_index, join_new


def _create_single_layer_type(base_type, layer):
    """Создание нового типа стены для одного слоя."""
    new_name = "{} | {}".format(base_type.Name, layer.name)

    # Проверяем, не существует ли уже такой тип
    collector = DB.FilteredElementCollector(revit.doc).OfClass(DB.WallType)
    for wt in collector:
        if wt.Name == new_name:
            return wt

    new_type = base_type.Duplicate(new_name)
    material_id = layer.material_id if layer.material_id else DB.ElementId.InvalidElementId

    # Создаём простую структуру с одним слоем
    structure = DB.CompoundStructure.CreateSimpleCompoundStructure(layer.width, material_id)
    structure.SetLayerFunction(0, layer.function)
    structure.SetNumberOfShellLayers(DB.ShellLayerType.Exterior, 0)
    structure.SetNumberOfShellLayers(DB.ShellLayerType.Interior, 0)

    new_type.SetCompoundStructure(structure)
    return new_type


def _copy_instance_parameters(source, target):
    """Копирование параметров между элементами."""
    # Пропускаем системные параметры
    skip_params = {
        DB.BuiltInParameter.WALL_KEY_REF_PARAM,
        DB.BuiltInParameter.WALL_HEIGHT_TYPE,
        DB.BuiltInParameter.WALL_USER_HEIGHT_PARAM,
        DB.BuiltInParameter.WALL_BASE_CONSTRAINT,
        DB.BuiltInParameter.WALL_TOP_CONSTRAINT,
        DB.BuiltInParameter.WALL_BASE_OFFSET,
        DB.BuiltInParameter.WALL_TOP_OFFSET,
    }

    for param in target.Parameters:
        if param.IsReadOnly:
            continue

        # Пропускаем системные параметры
        if param.Definition.BuiltInParameter in skip_params:
            continue

        src_param = source.LookupParameter(param.Definition.Name)
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
    """Создание новой стены для слоя с правильными параметрами."""
    location = base_wall.Location
    if not isinstance(location, DB.LocationCurve):
        raise ValueError("Стену с криволинейной геометрией обработать не удалось.")

    curve = location.Curve

    # Создаём или используем существующий тип стены
    if layer.wall_type is None:
        wall_type = _create_single_layer_type(base_type, layer)
        layer.wall_type = wall_type
    else:
        wall_type = layer.wall_type

    level_id = base_wall.LevelId

    # Получаем высоту оригинальной стены
    height_param = base_wall.get_Parameter(DB.BuiltInParameter.WALL_USER_HEIGHT_PARAM)
    height = height_param.AsDouble() if height_param else 3000.0 / 304.8  # По умолчанию 3м

    # Создаём новую стену с правильной высотой
    new_wall = DB.Wall.Create(
        revit.doc,
        curve,
        wall_type.Id,
        level_id,
        height,  # Используем реальную высоту
        0.0,  # Смещение от уровня
        base_wall.Flipped,
        False
    )

    # Копируем параметры высоты
    try:
        # Тип ограничения сверху
        top_constraint = base_wall.get_Parameter(DB.BuiltInParameter.WALL_HEIGHT_TYPE)
        if top_constraint:
            new_wall.get_Parameter(DB.BuiltInParameter.WALL_HEIGHT_TYPE).Set(top_constraint.AsInteger())

        # Верхнее ограничение
        top_level = base_wall.get_Parameter(DB.BuiltInParameter.WALL_TOP_CONSTRAINT)
        if top_level and top_level.AsElementId() != DB.ElementId.InvalidElementId:
            new_wall.get_Parameter(DB.BuiltInParameter.WALL_TOP_CONSTRAINT).Set(top_level.AsElementId())

        # Смещения
        base_offset = base_wall.get_Parameter(DB.BuiltInParameter.WALL_BASE_OFFSET)
        if base_offset:
            new_wall.get_Parameter(DB.BuiltInParameter.WALL_BASE_OFFSET).Set(base_offset.AsDouble())

        top_offset = base_wall.get_Parameter(DB.BuiltInParameter.WALL_TOP_OFFSET)
        if top_offset:
            new_wall.get_Parameter(DB.BuiltInParameter.WALL_TOP_OFFSET).Set(top_offset.AsDouble())
    except Exception:
        pass

    # Копируем остальные параметры
    _copy_instance_parameters(base_wall, new_wall)

    # Копируем структурное использование
    new_wall.StructuralUsage = base_wall.StructuralUsage

    # Вычисляем смещение для слоя
    orientation = base_wall.Orientation
    translation = orientation.Multiply(layer.offset_from_exterior)

    if translation.GetLength() > 1e-9:
        DB.ElementTransformUtils.MoveElement(revit.doc, new_wall.Id, translation)

    layer.translation = translation
    layer.new_wall = new_wall
    return new_wall


def _collect_hosted_instances(wall):
    """Сбор всех вставок (окна, двери) в стене."""
    collector = DB.FilteredElementCollector(revit.doc).OfClass(DB.FamilyInstance)
    hosted = []
    for inst in collector:
        try:
            host = inst.Host
            if host and host.Id == wall.Id:
                hosted.append(inst)
        except Exception:
            continue
    return hosted


def _create_hosted_instance(source_instance, host_wall, translation_vector):
    """Создание копии вставки (окна/двери) в новой стене."""
    location = source_instance.Location
    if not isinstance(location, DB.LocationPoint):
        return None

    point = location.Point + translation_vector
    symbol = source_instance.Symbol
    level_id = source_instance.LevelId

    # Активируем символ если нужно
    if not symbol.IsActive:
        symbol.Activate()
        revit.doc.Regenerate()

    try:
        new_instance = revit.doc.Create.NewFamilyInstance(
            point,
            symbol,
            host_wall,
            revit.doc.GetElement(level_id),
            DB.Structure.StructuralType.NonStructural,
        )
    except Exception:
        # Альтернативный метод создания
        try:
            new_instance = revit.doc.Create.NewFamilyInstance(
                point,
                symbol,
                host_wall,
                DB.Structure.StructuralType.NonStructural,
            )
        except Exception:
            return None

    # Копируем ориентацию
    new_location = new_instance.Location
    if isinstance(new_location, DB.LocationPoint):
        original_rotation = location.Rotation
        rotation_delta = original_rotation - new_location.Rotation
        if abs(rotation_delta) > 1e-6:
            axis = DB.Line.CreateBound(point, point + DB.XYZ.BasisZ)
            DB.ElementTransformUtils.RotateElement(revit.doc, new_instance.Id, axis, rotation_delta)

    # Копируем флипы
    try:
        if hasattr(new_instance, 'HandFlipped') and new_instance.HandFlipped != source_instance.HandFlipped:
            new_instance.flipHand()
    except Exception:
        pass

    try:
        if hasattr(new_instance, 'FacingFlipped') and new_instance.FacingFlipped != source_instance.FacingFlipped:
            new_instance.flipFacing()
    except Exception:
        pass

    _copy_instance_parameters(source_instance, new_instance)
    return new_instance


def _join_walls(walls):
    """Соединение геометрии стен между собой."""
    count = len(walls)
    joined_pairs = []

    for i in range(count):
        for j in range(i + 1, count):
            first = walls[i]
            second = walls[j]
            try:
                # Проверяем, не соединены ли уже
                if not DB.JoinGeometryUtils.AreElementsJoined(revit.doc, first, second):
                    DB.JoinGeometryUtils.JoinGeometry(revit.doc, first, second)
                    joined_pairs.append((i, j))
            except Exception:
                continue

    return len(joined_pairs)


def main():
    """Основная функция."""
    # Выбираем стену
    wall = _pick_wall()
    base_type = revit.doc.GetElement(wall.GetTypeId())

    # Получаем слои
    layers = _get_layers(wall)

    # Выбираем слои для разделения
    selected_layers = _select_layers(layers)

    # Собираем существующие типы стен
    wall_types = _collect_wall_types()

    # Настраиваем маппинг
    host_index, join_new = _ask_mapping(selected_layers, wall_types)

    # Собираем вставки
    hosted_instances = _collect_hosted_instances(wall)

    output = script.get_output()
    output.print_md("## Процесс разделения стены")

    with revit.Transaction("Разделить стену по слоям"):
        new_walls = []

        # Создаём новые стены
        output.print_md("### Создание новых стен:")
        for layer in selected_layers:
            try:
                new_wall = _create_new_wall(layer, wall, base_type)
                new_walls.append(new_wall)
                output.print_md("- ✓ Создана стена для слоя: **{}**".format(layer.display_name))
            except Exception as err:
                output.print_md("- ✗ Ошибка при создании стены для слоя {}: {}".format(layer.display_name, err))

        # Определяем слой для вставок
        host_layer = next((l for l in selected_layers if l.index == host_index), None)
        if host_layer is None and selected_layers:
            host_layer = selected_layers[0]

        # Получаем вектор смещения
        translation_vector = (
            host_layer.translation if host_layer and host_layer.translation else DB.XYZ.Zero
        )

        # Переносим вставки
        if hosted_instances:
            output.print_md("### Перенос окон и дверей:")
            transferred = 0
            for inst in hosted_instances:
                try:
                    if _create_hosted_instance(inst, host_layer.new_wall if host_layer else new_walls[0],
                                               translation_vector):
                        transferred += 1
                except Exception:
                    pass
            output.print_md("- Перенесено элементов: **{}** из **{}**".format(transferred, len(hosted_instances)))

        # Удаляем оригинальные вставки
        for inst in hosted_instances:
            try:
                revit.doc.Delete(inst.Id)
            except Exception:
                pass

        # Удаляем исходную стену
        try:
            revit.doc.Delete(wall.Id)
            output.print_md("### ✓ Исходная стена удалена")
        except Exception as err:
            output.print_md("### ✗ Не удалось удалить исходную стену: {}".format(err))

        # Соединяем новые стены
        if join_new and len(new_walls) > 1:
            joined = _join_walls(new_walls)
            if joined > 0:
                output.print_md("### ✓ Соединено пар стен: **{}**".format(joined))

    output.print_md("---")
    output.print_md("## **Готово!** Создано **{}** новых стен".format(len(new_walls)))


if __name__ == "__main__":
    main()