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
from System import MissingMemberException

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


def sanitize_name(name):
    """Удаление недопустимых символов и лишних пробелов из имени."""
    if not name:
        return ""

    invalid_chars = '<>:"/\\|?*'
    cleaned = []
    for char in name:
        if char in invalid_chars:
            cleaned.append('-')
        else:
            cleaned.append(char)

    normalized = "".join(cleaned)
    normalized = normalized.replace('\n', ' ').replace('\r', ' ')
    normalized = " ".join(normalized.split())
    return normalized.strip()


def generate_default_type_name(base_name, material_name, thickness_mm, index):
    """Формирование наглядного и уникального имени типа стены."""
    base = sanitize_name(base_name) or u"Стена"
    material = sanitize_name(material_name) or u"Слой {}".format(index + 1)
    thickness_value = int(round(thickness_mm or 0))
    return u"{}_L{}_{}_{}мм".format(base, index + 1, material, thickness_value)


class LayerUIData(object):
    """Описание слоя для отображения в пользовательском интерфейсе."""

    def __init__(self, layer_info, default_name):
        self.layer_info = layer_info
        self.MaterialName = layer_info.get('material_name') or u"Слой {}".format(layer_info.get('index', 0) + 1)
        self.ThicknessDisplay = u"{:.1f} мм".format(layer_info.get('width_mm', 0.0))
        self.ReplacementName = default_name
        self.DefaultName = default_name
        self.Selected = True

    @property
    def display_label(self):
        return u"{} ({:.1f} мм)".format(self.MaterialName, self.layer_info.get('width_mm', 0.0))


class LayerSelectionWindow(forms.WPFWindow):
    """Окно выбора слоёв и задания имён новых типов стен."""

    def __init__(self, layers_info, base_name):
        xaml_path = script.get_bundle_file('LayerSelection.xaml')
        forms.WPFWindow.__init__(self, xaml_path)

        self.layers_info = layers_info
        self.base_name = base_name
        self.layer_items = []
        self._selected_items = []

        self._build_layers()

    def _build_layers(self):
        """Создание элементов интерфейса для списка слоёв."""
        from System.Windows.Controls import CheckBox
        from System.Windows import Thickness

        for idx, layer in enumerate(self.layers_info):
            default_name = generate_default_type_name(
                self.base_name,
                layer.get('material_name'),
                layer.get('width_mm'),
                layer.get('index', idx)
            )

            item = LayerUIData(layer, default_name)
            self.layer_items.append(item)

            checkbox = CheckBox()
            checkbox.Content = item.display_label
            checkbox.IsChecked = True
            checkbox.Tag = item
            checkbox.Margin = Thickness(0, 0, 0, 6)
            checkbox.Checked += self._checkbox_changed
            checkbox.Unchecked += self._checkbox_changed
            self.LayersPanel.Children.Add(checkbox)

        self._refresh_grid()

    def _checkbox_changed(self, sender, args):
        item = getattr(sender, 'Tag', None)
        if item:
            item.Selected = bool(sender.IsChecked)
        self._refresh_grid()

    def _refresh_grid(self):
        selected = [item for item in self.layer_items if item.Selected]
        self.LayerGrid.ItemsSource = selected
        self.LayerGrid.Items.Refresh()

    def ok_click(self, sender, args):
        selected = [item for item in self.layer_items if item.Selected]
        if not selected:
            forms.alert(u"Выберите хотя бы один слой для разбивки.")
            return

        self._selected_items = selected
        self.DialogResult = True
        self.Close()

    def cancel_click(self, sender, args):
        self.DialogResult = False
        self.Close()

    def get_selected_layers(self):
        """Возвращает список выбранных слоёв с пользовательскими именами."""
        results = []
        for item in getattr(self, '_selected_items', []):
            layer_copy = dict(item.layer_info)
            custom_name = sanitize_name(item.ReplacementName) or item.DefaultName
            if not custom_name:
                custom_name = generate_default_type_name(
                    self.base_name,
                    layer_copy.get('material_name'),
                    layer_copy.get('width_mm'),
                    layer_copy.get('index', 0)
                )
            layer_copy['custom_type_name'] = custom_name
            results.append(layer_copy)
        return results

class WallLayerSeparator:
    def __init__(self):
        self.doc = doc
        self.uidoc = uidoc
        self.original_wall = None
        self.wall_type = None
        self.compound_structure = None
        self.new_walls = []
        self.hosted_elements = []

    def get_element_name(self, element, default="Без имени"):
        """Безопасное получение имени элемента Revit"""
        if not element:
            return default

        # Сначала пытаемся получить свойство Name напрямую
        try:
            name = element.Name
            if name:
                return name
        except (AttributeError, MissingMemberException):
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
            return default

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
                material_name = self.get_element_name(material, default="Без материала")

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
                    'material_name': self.get_element_name(material, default="Основной материал"),
                    'width_mm': round(total_width * 304.8, 1)
                })

        return layers_info

    def get_creation_height(self):
        """Получение высоты исходной стены для создания новых"""

        # Пробуем получить «несвязанную» высоту стены
        try:
            height_param = self.original_wall.get_Parameter(
                BuiltInParameter.WALL_USER_HEIGHT_PARAM
            )
            if height_param and height_param.HasValue:
                height = height_param.AsDouble()
                if height and height > 0:
                    return height
        except Exception:
            pass

        # Если параметр не дал результата — оцениваем высоту по габаритному контейнеру
        try:
            bbox = self.original_wall.get_BoundingBox(None)
            if bbox:
                height = bbox.Max.Z - bbox.Min.Z
                if height and height > 0:
                    return height
        except Exception:
            pass

        # Последний вариант — используем стандартное значение 3000 мм
        return UnitUtils.ConvertToInternalUnits(3000, UnitTypeId.Millimeters)

    def create_single_layer_wall_type(self, layer_info, base_name):
        """Создание нового типа стены с одним слоем"""

        desired_name = layer_info.get('custom_type_name')
        if desired_name:
            new_name = sanitize_name(desired_name)
        else:
            new_name = ""

        if not new_name:
            new_name = generate_default_type_name(
                base_name,
                layer_info.get('material_name'),
                layer_info.get('width_mm'),
                layer_info.get('index', 0)
            )

        # Ограничиваем длину имени, чтобы избежать ошибок API
        if len(new_name) > 60:
            new_name = new_name[:60]

        # Проверяем, существует ли уже такой тип
        existing = None
        collector = FilteredElementCollector(doc).OfClass(WallType)
        for wt in collector:
            wt_name_param = wt.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
            wt_name = wt_name_param.AsString() if wt_name_param else None
            if wt_name == new_name:
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

        # Собираем кандидатов из всех экземпляров семейств в документе
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

        # Список параметров для копирования. Для каждого имени указаны варианты,
        # которые встречаются в разных версиях Revit API. Это защищает код от
        # AttributeError при отсутствии конкретного перечислителя.
        param_name_map = [
            ("WALL_BASE_CONSTRAINT", "WALL_BASE_CONSTRAINT_PARAM"),
            ("WALL_BASE_OFFSET", "WALL_BASE_OFFSET_PARAM"),
            ("WALL_HEIGHT_TYPE", "WALL_TOP_CONSTRAINT", "WALL_TOP_CONSTRAINT_PARAM"),
            ("WALL_TOP_CONSTRAINT", "WALL_TOP_CONSTRAINT_PARAM", "WALL_HEIGHT_TYPE"),
            ("WALL_TOP_OFFSET", "WALL_TOP_OFFSET_PARAM"),
            ("WALL_USER_HEIGHT_PARAM", "WALL_USER_HEIGHT")
        ]

        params_to_copy = []
        resolved_params = set()

        for names in param_name_map:
            if isinstance(names, str):
                candidates = (names,)
            else:
                candidates = names

            resolved_param = None
            for name in candidates:
                param_id = getattr(BuiltInParameter, name, None)
                if param_id is not None:
                    resolved_param = param_id
                    if param_id not in resolved_params:
                        params_to_copy.append(param_id)
                        resolved_params.add(param_id)
                    break

            if resolved_param is None:
                try:
                    output.print_md(
                        u"⚠️ Параметр `{}` недоступен в данной версии API".format(candidates[0])
                    )
                except Exception:
                    pass

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

    def show_layer_selection_dialog(self, layers_info, base_name):
        """Отображение окна выбора слоёв и возврат выбранных данных."""

        try:
            window = LayerSelectionWindow(layers_info, base_name)
            dialog_result = window.ShowDialog()
        except Exception as e:
            output.print_md(
                u"⚠️ Не удалось отобразить окно выбора слоёв. Используются все слои.\nПричина: {}".format(e)
            )
            return layers_info

        if dialog_result is None:
            return None

        try:
            user_confirmed = bool(dialog_result)
        except Exception:
            user_confirmed = False

        if user_confirmed:
            selected = window.get_selected_layers()
            if selected:
                return selected
            return None

        return None

    def execute(self):
        """Основной метод выполнения"""

        # 1. Выбор стены
        if not self.select_wall():
            return

        wall_type_name = self.get_element_name(self.wall_type)
        output.print_md("## Разбивка стены: **{}**".format(wall_type_name))

        # 2. Получение информации о слоях
        layers_info = self.get_wall_info()

        output.print_md("### Найдено слоёв: **{}**".format(len(layers_info)))
        for layer in layers_info:
            output.print_md("- **{}**: {} мм".format(
                layer['material_name'],
                layer['width_mm']
            ))

        base_name = sanitize_name(wall_type_name) or "Стена"

        # 3. Открытие окна выбора слоёв
        selected_layers = self.show_layer_selection_dialog(layers_info, base_name)
        if selected_layers is None:
            output.print_md("⚠️ Разбивка отменена пользователем.")
            return

        layers_info = selected_layers

        output.print_md("### К обработке выбрано слоёв: **{}**".format(len(layers_info)))
        for layer in layers_info:
            custom_name = layer.get('custom_type_name')
            if not custom_name:
                custom_name = generate_default_type_name(
                    base_name,
                    layer.get('material_name'),
                    layer.get('width_mm'),
                    layer.get('index', 0)
                )
            output.print_md("- **{}** → новый тип: `{}`".format(
                layer['material_name'],
                custom_name
            ))

        # 4. Получение вложенных элементов
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

        # 5. Начинаем транзакцию
        with revit.Transaction("Разбивка стены на слои"):

            # 6. Создаём новые типы стен
            output.print_md("### Создание новых типов стен...")
            new_wall_types = []

            for layer in layers_info:
                new_type = self.create_single_layer_wall_type(layer, base_name)
                new_wall_types.append(new_type)
                layer['new_type'] = new_type

            # 7. Рассчитываем позиции новых стен
            positions = self.calculate_wall_positions(layers_info)

            # 8. Создаём новые стены
            output.print_md("### Создание новых стен...")
            for pos_data in positions:
                new_wall = Wall.Create(
                    doc,
                    pos_data['curve'],
                    pos_data['layer']['new_type'].Id,
                    self.original_wall.LevelId,
                    self.get_creation_height(),
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

            # 9. Переносим вложенные элементы на ближайшую новую стену
            if self.hosted_elements and self.new_walls:
                output.print_md("### Перенос вложенных элементов...")

                def _wall_width(target_wall):
                    try:
                        structure = target_wall.WallType.GetCompoundStructure()
                        if structure and structure.LayerCount:
                            return structure.GetLayerWidth(0)
                    except Exception:
                        pass

                    try:
                        width_param = target_wall.WallType.get_Parameter(
                            BuiltInParameter.WALL_ATTR_WIDTH_PARAM
                        )
                        if width_param and width_param.HasValue:
                            return width_param.AsDouble()
                    except Exception:
                        pass

                    return 0

                # Находим стену с максимальной толщиной (обычно несущая)
                main_wall = max(self.new_walls, key=_wall_width)

                for element in self.hosted_elements:
                    try:
                        element.Host = main_wall
                        output.print_md("✓ Перенесён элемент: **{}**".format(
                            self.get_element_name(element)
                        ))
                    except Exception as e:
                        output.print_md("✗ Ошибка переноса: {}".format(str(e)))

            # 10. Соединяем новые стены между собой
            output.print_md("### Соединение стен...")
            for i, wall1 in enumerate(self.new_walls):
                for wall2 in self.new_walls[i + 1:]:
                    try:
                        JoinGeometryUtils.JoinGeometry(doc, wall1, wall2)
                    except:
                        pass  # Игнорируем ошибки соединения

            # 11. Удаляем исходную стену
            doc.Delete(self.original_wall.Id)
            output.print_md("### ✓ Исходная стена удалена")

        output.print_md("## ✅ Разбивка завершена успешно!")
        output.print_md("Создано стен: **{}**".format(len(self.new_walls)))


# Запуск
if __name__ == '__main__':
    separator = WallLayerSeparator()
    separator.execute()
