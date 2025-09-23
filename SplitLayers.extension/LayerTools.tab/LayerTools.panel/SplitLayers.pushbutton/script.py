# -*- coding: utf-8 -*-
"""
Разбивка многослойной стены на отдельные стены по слоям
для PyRevit / Revit 2022
"""

__title__ = 'Разбить\nСтену'
__author__ = 'Wall Layers Separator'

import clr
import sys
import math

# Импорт Revit API
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.DB.Structure import *

# Импорт PyRevit
from pyrevit import revit, DB, UI
from pyrevit import forms
from pyrevit import script

# Получаем документ и UI
doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()


class WallLayerSeparator:
    def __init__(self):
        self.doc = doc
        self.uidoc = uidoc
        self.original_wall = None
        self.wall_type = None
        self.compound_structure = None
        self.new_walls = []
        self.hosted_elements = []

    def get_element_name(self, element):
        """Безопасное получение имени элемента Revit"""
        if not element:
            return "Без имени"

        # Сначала пытаемся получить свойство Name напрямую
        try:
            name = element.Name
            if name:
                return name
        except AttributeError:
            pass

        # Если свойства нет, пробуем получить значение из параметров
        param_ids = [
            BuiltInParameter.ALL_MODEL_TYPE_NAME,
            BuiltInParameter.SYMBOL_NAME_PARAM,
            BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM
        ]

        for param_id in param_ids:
            try:
                param = element.get_Parameter(param_id)
                if param and param.HasValue:
                    name = param.AsString()
                    if name:
                        return name
            except Exception:
                continue

        # В крайнем случае возвращаем идентификатор элемента
        try:
            return "ID {}".format(element.Id.IntegerValue)
        except Exception:
            return "Без имени"

    def select_wall(self):
        """Выбор стены пользователем"""
        try:
            reference = uidoc.Selection.PickObject(
                ObjectType.Element,
                "Выберите многослойную стену для разбивки"
            )
            self.original_wall = doc.GetElement(reference.ElementId)

            if not isinstance(self.original_wall, Wall):
                forms.alert("Выбран не стеновой элемент!", exitscript=True)
                return False

            self.wall_type = self.original_wall.WallType

            # Пробуем получить структуру разными способами
            self.compound_structure = None

            # Способ 1: напрямую от типа
            try:
                self.compound_structure = self.wall_type.GetCompoundStructure()
            except:
                pass

            # Способ 2: через параметр WALL_STRUCTURE_ID_PARAM
            if not self.compound_structure:
                try:
                    struct_param = self.wall_type.get_Parameter(
                        BuiltInParameter.WALL_STRUCTURE_ID_PARAM
                    )
                    if struct_param:
                        struct_id = struct_param.AsElementId()
                        if struct_id and struct_id != ElementId.InvalidElementId:
                            struct_elem = doc.GetElement(struct_id)
                            if hasattr(struct_elem, 'GetCompoundStructure'):
                                self.compound_structure = struct_elem.GetCompoundStructure()
                except:
                    pass

            # Способ 3: проверяем, может это системная стена
            if not self.compound_structure:
                try:
                    # Для системных семейств стен
                    if self.wall_type.Kind == WallKind.Basic:
                        # Пробуем создать временную структуру
                        width_param = self.wall_type.get_Parameter(
                            BuiltInParameter.WALL_ATTR_WIDTH_PARAM
                        )
                        if width_param:
                            width = width_param.AsDouble()
                            if width > 0:
                                # Это однослойная стена - создаём структуру
                                material_param = self.wall_type.get_Parameter(
                                    BuiltInParameter.STRUCTURAL_MATERIAL_PARAM
                                )
                                material_id = material_param.AsElementId() if material_param else ElementId.InvalidElementId

                                # Проверяем, действительно ли это однослойная стена
                                output.print_md("⚠️ Обнаружена базовая стена без составной структуры")
                                output.print_md("Тип: {}".format(self.get_element_name(self.wall_type)))
                                output.print_md("Толщина: {} мм".format(round(width * 304.8, 1)))

                                forms.alert("Это базовая однослойная стена!\nИспользуйте для многослойных стен.",
                                            exitscript=True)
                                return False
                except:
                    pass

            # Проверяем на наличие редактированного профиля
            if not self.compound_structure:
                # Проверка на модифицированную геометрию
                try:
                    if hasattr(self.original_wall, 'IsModified') and self.original_wall.IsModified:
                        output.print_md("⚠️ **Стена имеет модифицированную геометрию**")
                except:
                    pass

                # Проверка на вертикальную структуру
                has_profile = False
                try:
                    profile_param = self.original_wall.get_Parameter(
                        BuiltInParameter.WALL_SWEEP_PROFILE_PARAM
                    )
                    if profile_param and profile_param.HasValue:
                        has_profile = True
                except:
                    pass

                if has_profile:
                    output.print_md("⚠️ **Стена имеет редактируемый профиль**")

            if not self.compound_structure:
                # Выводим отладочную информацию
                output.print_md("### Отладочная информация:")
                output.print_md("- Имя типа: **{}**".format(self.get_element_name(self.wall_type)))
                output.print_md("- ID типа: **{}**".format(self.wall_type.Id))
                output.print_md("- Тип стены (Kind): **{}**".format(self.wall_type.Kind))

                # Проверяем все параметры типа
                params_info = []
                for param in self.wall_type.Parameters:
                    if param.HasValue:
                        try:
                            value = param.AsValueString() or str(param.AsElementId().IntegerValue)
                            params_info.append("{}: {}".format(param.Definition.Name, value))
                        except:
                            pass

                if params_info:
                    output.print_md("### Параметры типа стены:")
                    for info in params_info[:10]:  # Первые 10 параметров
                        output.print_md("- {}".format(info))

                forms.alert(
                    "Стена не имеет многослойной структуры!\n\nВозможные причины:\n1. Это не составная стена\n2. Стена имеет модификации\n3. Используется сложный профиль",
                    exitscript=True)
                return False

            if self.compound_structure.LayerCount < 2:
                forms.alert("Стена имеет только {} слой!\nДля разбивки нужно минимум 2 слоя.".format(
                    self.compound_structure.LayerCount
                ), exitscript=True)
                return False

            return True

        except Exception as e:
            forms.alert("Ошибка при выборе стены:\n{}".format(str(e)), exitscript=True)
            return False

    def get_wall_info(self):
        """Получение информации о слоях стены"""
        layers_info = []

        # Если есть CompoundStructure - используем её
        if self.compound_structure:
            for i in range(self.compound_structure.LayerCount):
                layer_func = self.compound_structure.GetLayerFunction(i)
                layer_width = self.compound_structure.GetLayerWidth(i)
                layer_material_id = self.compound_structure.GetMaterialId(i)

                material = doc.GetElement(layer_material_id) if layer_material_id else None
                material_name = material.Name if material else "Без материала"

                layers_info.append({
                    'index': i,
                    'function': layer_func,
                    'width': layer_width,
                    'material_id': layer_material_id,
                    'material_name': material_name,
                    'width_mm': round(layer_width * 304.8, 1)  # футы в мм
                })
        else:
            # Альтернативный метод для стен без CompoundStructure
            # Пытаемся получить информацию через параметры
            output.print_md("⚠️ Используется альтернативный метод получения слоёв")

            # Получаем общую толщину
            width_param = self.wall_type.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM)
            if width_param:
                total_width = width_param.AsDouble()

                # Создаём псевдо-слой
                material_param = self.wall_type.get_Parameter(BuiltInParameter.STRUCTURAL_MATERIAL_PARAM)
                material_id = material_param.AsElementId() if material_param else ElementId.InvalidElementId
                material = doc.GetElement(material_id) if material_id != ElementId.InvalidElementId else None

                layers_info.append({
                    'index': 0,
                    'function': MaterialFunctionAssignment.Structure,
                    'width': total_width,
                    'material_id': material_id,
                    'material_name': material.Name if material else "Основной материал",
                    'width_mm': round(total_width * 304.8, 1)
                })

        return layers_info

    def create_single_layer_wall_type(self, layer_info, base_name):
        """Создание нового типа стены с одним слоем"""

        # Генерируем уникальное имя
        new_name = "{}_{}_{}мм".format(
            base_name,
            layer_info['material_name'][:20],  # Ограничиваем длину
            int(layer_info['width_mm'])
        )

        # Проверяем, существует ли уже такой тип
        existing = None
        collector = FilteredElementCollector(doc).OfClass(WallType)
        for wt in collector:
            if wt.Name == new_name:
                existing = wt
                break

        if existing:
            return existing

        # Создаём новый тип
        new_wall_type = self.wall_type.Duplicate(new_name)

        # Создаём структуру с одним слоем
        try:
            new_compound = CompoundStructure.CreateSingleLayerCompoundStructure(
                layer_info['function'],
                layer_info['width'],
                layer_info['material_id']
            )

            # Применяем структуру к новому типу
            new_wall_type.SetCompoundStructure(new_compound)
        except Exception as e:
            # Если не получилось создать CompoundStructure, просто меняем толщину
            output.print_md("⚠️ Альтернативное создание типа: {}".format(new_name))
            width_param = new_wall_type.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM)
            if width_param and not width_param.IsReadOnly:
                width_param.Set(layer_info['width'])

        return new_wall_type

    def calculate_wall_positions(self, layers_info):
        """Расчёт позиций для новых стен"""

        # Получаем линию расположения стены
        location_curve = self.original_wall.Location
        if not isinstance(location_curve, LocationCurve):
            raise Exception("Стена не имеет линии расположения")

        wall_line = location_curve.Curve

        # Получаем вектор перпендикулярный стене
        start_point = wall_line.GetEndPoint(0)
        end_point = wall_line.GetEndPoint(1)

        # Вектор вдоль стены
        wall_vector = (end_point - start_point).Normalize()

        # Перпендикулярный вектор (в плане XY)
        perpendicular = XYZ(-wall_vector.Y, wall_vector.X, 0)

        # Определяем сторону ориентации стены
        if self.original_wall.Flipped:
            perpendicular = -perpendicular

        # Получаем общую толщину стены
        total_width = sum([l['width'] for l in layers_info])

        # Рассчитываем смещения для каждого слоя
        positions = []
        current_offset = -total_width / 2.0  # Начинаем от внешней грани

        for layer in layers_info:
            # Центр текущего слоя
            layer_center_offset = current_offset + layer['width'] / 2.0

            # Создаём линию для новой стены
            offset_vector = perpendicular * layer_center_offset
            new_start = start_point + offset_vector
            new_end = end_point + offset_vector

            new_line = Line.CreateBound(new_start, new_end)

            positions.append({
                'layer': layer,
                'curve': new_line,
                'offset': layer_center_offset
            })

            current_offset += layer['width']

        return positions

    def get_hosted_elements(self):
        """Получение всех вложенных элементов (окна, двери и т.д.)"""

        # Используем фильтр для поиска элементов, привязанных к стене
        filter = FamilyInstanceFilter(doc, self.original_wall.Id)
        collector = FilteredElementCollector(doc)

        # Собираем все FamilyInstance элементы
        all_instances = collector.OfClass(FamilyInstance).ToElements()

        hosted = []
        for instance in all_instances:
            if instance.Host and instance.Host.Id == self.original_wall.Id:
                hosted.append(instance)

        return hosted

    def copy_parameters(self, from_wall, to_wall):
        """Копирование параметров со старой стены на новую"""

        # Список параметров для копирования
        params_to_copy = [
            BuiltInParameter.WALL_BASE_CONSTRAINT,
            BuiltInParameter.WALL_BASE_OFFSET,
            BuiltInParameter.WALL_HEIGHT_TYPE,
            BuiltInParameter.WALL_TOP_CONSTRAINT,
            BuiltInParameter.WALL_TOP_OFFSET,
            BuiltInParameter.WALL_USER_HEIGHT_PARAM
        ]

        for param_id in params_to_copy:
            try:
                from_param = from_wall.get_Parameter(param_id)
                to_param = to_wall.get_Parameter(param_id)

                if from_param and to_param and not to_param.IsReadOnly:
                    if from_param.StorageType == StorageType.Double:
                        to_param.Set(from_param.AsDouble())
                    elif from_param.StorageType == StorageType.Integer:
                        to_param.Set(from_param.AsInteger())
                    elif from_param.StorageType == StorageType.ElementId:
                        to_param.Set(from_param.AsElementId())
                    elif from_param.StorageType == StorageType.String:
                        to_param.Set(from_param.AsString())
            except:
                continue

    def execute(self):
        """Основной метод выполнения"""

        # 1. Выбор стены
        if not self.select_wall():
            return

        output.print_md("## Разбивка стены: **{}**".format(self.get_element_name(self.wall_type)))

        # 2. Получение информации о слоях
        layers_info = self.get_wall_info()

        output.print_md("### Найдено слоёв: **{}**".format(len(layers_info)))
        for layer in layers_info:
            output.print_md("- **{}**: {} мм".format(
                layer['material_name'],
                layer['width_mm']
            ))

        # 3. Получение вложенных элементов
        self.hosted_elements = self.get_hosted_elements()
        if self.hosted_elements:
            output.print_md("### Найдено вложенных элементов: **{}**".format(
                len(self.hosted_elements)
            ))

        # Запрос подтверждения
        if not forms.alert(
                "Разбить стену на {} отдельных слоёв?".format(len(layers_info)),
                yes=True, no=True
        ):
            return

        # 4. Начинаем транзакцию
        with revit.Transaction("Разбивка стены на слои"):

            # 5. Создаём новые типы стен
            output.print_md("### Создание новых типов стен...")
            new_wall_types = []
            wall_type_name = self.get_element_name(self.wall_type)
            base_name = wall_type_name.split('_')[0] if wall_type_name else "Стена"

            for layer in layers_info:
                new_type = self.create_single_layer_wall_type(layer, base_name)
                new_wall_types.append(new_type)
                layer['new_type'] = new_type

            # 6. Рассчитываем позиции новых стен
            positions = self.calculate_wall_positions(layers_info)

            # 7. Создаём новые стены
            output.print_md("### Создание новых стен...")
            for pos_data in positions:
                new_wall = Wall.Create(
                    doc,
                    pos_data['curve'],
                    pos_data['layer']['new_type'].Id,
                    self.original_wall.LevelId,
                    self.original_wall.Height,
                    0,  # offset
                    self.original_wall.Flipped,
                    False  # structural
                )

                # Копируем параметры
                self.copy_parameters(self.original_wall, new_wall)

                self.new_walls.append(new_wall)

                output.print_md("✓ Создана стена: **{}**".format(
                    self.get_element_name(pos_data['layer']['new_type'])
                ))

            # 8. Переносим вложенные элементы на ближайшую новую стену
            if self.hosted_elements and self.new_walls:
                output.print_md("### Перенос вложенных элементов...")

                # Находим стену с максимальной толщиной (обычно несущая)
                main_wall = max(self.new_walls,
                                key=lambda w: w.WallType.GetCompoundStructure().GetLayerWidth(0))

                for element in self.hosted_elements:
                    try:
                        element.Host = main_wall
                        output.print_md("✓ Перенесён элемент: **{}**".format(
                            self.get_element_name(element)
                        ))
                    except Exception as e:
                        output.print_md("✗ Ошибка переноса: {}".format(str(e)))

            # 9. Соединяем новые стены между собой
            output.print_md("### Соединение стен...")
            for i, wall1 in enumerate(self.new_walls):
                for wall2 in self.new_walls[i + 1:]:
                    try:
                        JoinGeometryUtils.JoinGeometry(doc, wall1, wall2)
                    except:
                        pass  # Игнорируем ошибки соединения

            # 10. Удаляем исходную стену
            doc.Delete(self.original_wall.Id)
            output.print_md("### ✓ Исходная стена удалена")

        output.print_md("## ✅ Разбивка завершена успешно!")
        output.print_md("Создано стен: **{}**".format(len(self.new_walls)))


# Запуск
if __name__ == '__main__':
    separator = WallLayerSeparator()
    separator.execute()
