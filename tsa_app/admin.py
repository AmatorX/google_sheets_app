
import logging
from time import sleep
from datetime import date
from django.contrib import admin, messages
from django.db import IntegrityError

from sheets.photos_sheet import PhotosTable
# from django_celery_beat.models import PeriodicTask, IntervalSchedule, CrontabSchedule, SolarSchedule, ClockedSchedule

from sheets.sheet1 import Sheet1Table
from sheets.work_time_sheet import WorkTimeTable
from .models import Worker, BuildObject, Material, Tool, ToolsSheet, WorkEntry, BuildBudgetHistory, ForemanAndWorkersKPISheet
from django.contrib.admin import SimpleListFilter

from django.contrib import admin
from .models import BuildObject


logger = logging.getLogger(__name__)


class WorkEntryAdmin(admin.ModelAdmin):
    list_display = ('id', 'worker', 'build_object', 'worked_hours', 'date', 'short_materials')
    list_filter = ('worker', 'build_object', 'date')
    search_fields = ('worker__name', 'build_object__name')
    date_hierarchy = 'date'
    ordering = ('-date',)

    def short_materials(self, obj):
        """Краткое представление использованных материалов."""
        try:
            materials = obj.get_materials_used()
            return ', '.join([f"{k}: {v}" for k, v in list(materials.items())[:3]]) + ('...' if len(materials) > 3 else '')
        except Exception:
            return 'Ошибка разбора'
    short_materials.short_description = 'Materials'

class WorkerAdmin(admin.ModelAdmin):
    list_display = ('name', 'tg_id', 'build_obj', 'salary', 'title', 'foreman')
    list_display_links = ('name',)
    list_filter = ('build_obj', 'salary', 'title')
    search_fields = ('name', 'tg_id', 'build_obj')
    list_editable = ('build_obj',)
    fieldsets = (
        (None, {
            'fields': ('name', 'tg_id', 'salary', 'foreman', 'build_obj', 'phone_number', 'email')
        }),
        ('Advanced options', {
            'classes': ('collapse',),
            'fields': ('employment_agreement', 'over_time', 'start_to_work', 'title', 'payroll_eligible', 'payroll', 'resign_agreement', 'benefits_eligible', 'birthday', 'issued', 'expiry', 'address', 'tickets_available'),
        }),
    )
    list_per_page = 20

    def save_model(self, request, obj, form, change):
        super().save_model(request, obj, form, change)
        if not change:  # Проверяем, что объект был создан, а не обновлен
            pass
            # append_user_names_to_tables(obj)  # функция для записи данных
            # Отображаем сообщение в админке
            # messages.success(request, 'Данные записаны успешно!')


class MaterialAdmin(admin.ModelAdmin):
    fields = ('name', 'price')
    list_display = ('name', 'price')


class BuildObjectFilter(SimpleListFilter):
    title = 'Build Object'  # Название фильтра в админке
    parameter_name = 'build_obj'  # Параметр фильтра

    def lookups(self, request, model_admin):
        # Возвращаем список опций для фильтра
        build_objs = set(obj.assigned_to.build_obj for obj in model_admin.model.objects.all() if obj.assigned_to and obj.assigned_to.build_obj)
        return [(obj.id, obj.name) for obj in build_objs]

    def queryset(self, request, queryset):
        # Фильтруем queryset в зависимости от выбранного значения
        if self.value():
            return queryset.filter(assigned_to__build_obj__id=self.value())
        return queryset


class ToolAdmin(admin.ModelAdmin):
    fields = ('name', 'tool_id', 'price', 'date_of_issue', 'assigned_to', 'tools_sheet')
    list_display = ('name', 'tool_id', 'price', 'date_of_issue', 'assigned_to', 'get_build_obj_name')
    list_editable = ('assigned_to', 'price', 'date_of_issue',)
    list_filter = (BuildObjectFilter, 'name', 'tool_id', 'date_of_issue', 'assigned_to', )
    list_per_page = 50

    def get_build_obj_name(self, obj):
        return obj.assigned_to.build_obj.name if obj.assigned_to and obj.assigned_to.build_obj else 'N/A'

    get_build_obj_name.short_description = 'Build Object'


class ToolsSheetAdmin(admin.ModelAdmin):
    fields = ('name', 'sh_url')



@admin.action(description="Write a summary of the materials consumption in the table")
def write_materials_summary_action(modeladmin, request, queryset):
    for obj in queryset:
        try:
            table = Sheet1Table(obj)
            table.write_materials_summary()
        except Exception as e:
            modeladmin.message_user(request, f"Ошибка при обработке '{obj.name}': {e}", level="error")
        else:
            modeladmin.message_user(request, f"Сводка материалов записана в таблицу для '{obj.name}'")

@admin.action(description="Write summary data on the real average earnings per hour for employees")
def write_users_kpi_data(modeladmin, request, queryset):
    for obj in queryset:
        try:
            table = Sheet1Table(obj)
            table.write_summary_workers()
        except Exception as e:
            modeladmin.message_user(request, f"Ошибка при обработке '{obj.name}': {e}", level="error")
        else:
            modeladmin.message_user(request, f"Сводка результатов работник записана в таблицу для '{obj.name}'")


@admin.action(description="Write Work Time tables")
def update_all_worktime_tables(modeladmin, request, queryset):
    logger.info("Запущено обновление WorkTime таблиц")

    # build_objects = BuildObject.objects.all()

    for obj in queryset:
        try:
            logger.info(f"Обновляем WorkTime таблицу для объекта: {obj.name}")
            table = WorkTimeTable(obj=obj)
            table.update_current_chunk()
        except Exception as e:
            logger.error(f"Ошибка при обновлении объекта {obj.name}: {e}")
        sleep(10)
    messages.success(request, f"Обновление WorkTime завершено для {queryset.count()} объекта(ов)")

    logger.info("Обновление work time таблиц завершено")


@admin.action(description="Record monthly summary data")
def record_summary_data(modeladmin, request, queryset):
    logger.info("Запущена запись сводных данных за месяц")
    for obj in queryset:
        try:
            logger.info(f"Запись сводной таблицы за месяц для объекта: {obj.name}")
            table = Sheet1Table(obj=obj)
            table.write_summary()
        except Exception as e:
            logger.error(f"Ошибка при обновлении объекта {obj.name}: {e}")
        sleep(10)

    logger.info("Обновление сводных таблиц завершено")


@admin.action(description="Update tables for photo reports")
def update_photos_tables(modeladmin, request, queryset):
    logger.info("Запуск обновления таблиц фотоотчётов")

    for obj in queryset:
        try:
            logger.info(f"Обрабатываем объект: {obj.name}")
            photos_table = PhotosTable(obj=obj)
            photos_table.write_missing_worker_tables()
        except Exception as e:
            logger.error(f"Ошибка при обработке объекта {obj.name}: {e}")
        sleep(10)

    logger.info("Обновление фото-таблиц завершено")

@admin.action(description="Update the result tables for sheet1")
def run_monthly_summary_tasks(modeladmin, request, queryset):

    logger.info("Сегодня последний день месяца по времени Калгари. Запуск записи сводных данных...")
    all_workers_data = []
    for obj in queryset:
        try:
            # 1. KPI по пользователям
            table = Sheet1Table(obj)
            workers_data = table.write_summary_workers()
            logger.info(f"Сводка результатов работников записана для '{obj.name}'")
            all_workers_data.extend(workers_data)
            sleep(5)
            # 2. Сводка по объекту
            table.write_summary()
            logger.info(f"Сводная таблица за месяц записана для '{obj.name}'")
            sleep(5)

        except Exception as e:
            logger.error(f"Ошибка при обработке '{obj.name}': {e}")
        logger.info(f"Таблицы для объекта {obj.name} записааны, ожидание 60 сек ")

    logger.info("Месячная задача завершена.")


class BuildObjectAdmin(admin.ModelAdmin):
    actions = [update_all_worktime_tables, write_materials_summary_action, write_users_kpi_data, record_summary_data, update_photos_tables, run_monthly_summary_tasks]
    fields = ('name', 'total_budget', 'material', 'current_budget', 'sh_url', 'archive')
    list_display = ('name', 'archive', 'total_budget', 'current_budget', 'display_materials')
    list_filter = ['archive']
    list_editable = ('archive',)

    def display_materials(self, obj):
        return ", ".join([material.name for material in obj.material.all()])
    display_materials.short_description = 'Materials'

    # Запрещаем удаление объекта
    def has_delete_permission(self, request, obj=None):
        return False



class BuildBudgetHistoryAdmin(admin.ModelAdmin):
    list_display = ('build_object', 'date', 'current_budget')  # Отображаемые поля в списке
    list_filter = ('build_object', 'date')  # Фильтры справа
    search_fields = ('build_object__name',)  # Поиск по имени объекта
    date_hierarchy = 'date'  # навигация по датам сверху
    ordering = ('-date',)

@admin.register(ForemanAndWorkersKPISheet)
class ForemanAndWorkersKPIAdmin(admin.ModelAdmin):
    list_display = ('name', 'sh_url', 'created_at')
    search_fields = ('name',)


admin.site.register(Worker, WorkerAdmin)
admin.site.register(BuildBudgetHistory, BuildBudgetHistoryAdmin)
admin.site.register(BuildObject, BuildObjectAdmin)
admin.site.register(Material, MaterialAdmin)
admin.site.register(ToolsSheet, ToolsSheetAdmin)
admin.site.register(Tool, ToolAdmin)
admin.site.register(WorkEntry, WorkEntryAdmin)

