
import logging
from time import sleep
from datetime import date
from django.contrib import admin, messages
from django.shortcuts import redirect, render
from django.contrib.admin.helpers import ACTION_CHECKBOX_NAME

from sheets.photos_sheet import PhotosTable
from sheets.sheet1 import Sheet1Table
from sheets.tasks import update_general_statistic_task
from sheets.tools_sheet import ToolsTable
from sheets.work_time_sheet import WorkTimeTable
from .form import MonthYearForm
from .models import Worker, BuildObject, Material, Tool, ToolsSheet, WorkEntry, BuildBudgetHistory, ForemanAndWorkersKPISheet, MediaProxy, GeneralStatisticSheet, WorkSpecialization
from django.contrib.admin import SimpleListFilter
from django.contrib import admin


logger = logging.getLogger(__name__)


class MediaBrowserAdmin(admin.ModelAdmin):
    def changelist_view(self, request, extra_context=None):
        return redirect('media_browser')

    def has_add_permission(self, request): return False
    def has_change_permission(self, request, obj=None): return False
    def has_delete_permission(self, request, obj=None): return False


def with_month_year_form(template_name="admin/month_year_form.html"):
    """
    Декоратор для actions в админке, чтобы перед показом формы запросить у пользователя месяц и год
    и передать эти значения в оригинальную функцию действия

    :param template_name: Путь к HTML-шаблону формы, по умолчанию 'admin/month_year_form.html'
    :return: Обёрнутая admin action функция, с дополнительными параметрами month и year
    """
    def decorator(action_func):
        def wrapper(modeladmin, request, queryset):
            if "apply" in request.POST:
                form = MonthYearForm(request.POST)
                if form.is_valid():
                    month = form.cleaned_data["month"]
                    year = form.cleaned_data["year"]
                    ids = request.POST.getlist("_selected_action")
                    selected_objects = queryset.model.objects.filter(pk__in=ids)
                    return action_func(modeladmin, request, selected_objects, month, year)
            else:
                form = MonthYearForm(initial={
                    "_selected_action": request.POST.getlist(ACTION_CHECKBOX_NAME)
                })

            return render(request, template_name, {
                "objects": queryset,
                "form": form,
                "title": "Select the month and year to count and record the data.",
                "action_checkbox_name": ACTION_CHECKBOX_NAME,
                "action_name": action_func.__name__,
            })

        wrapper.__name__ = action_func.__name__
        wrapper.__doc__ = action_func.__doc__
        return wrapper
    return decorator


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


class ArchiveFilter(admin.SimpleListFilter):
    title = 'archive'
    parameter_name = 'is_archived'

    def lookups(self, request, model_admin):
        return [
            ('no', 'Not archive'),
            ('yes', 'Archive'),
        ]

    def queryset(self, request, queryset):
        if self.value() == 'no':
            return queryset.filter(is_archived=False)
        elif self.value() == 'yes':
            return queryset.filter(is_archived=True)
        return queryset  # "all"


class WorkerAdmin(admin.ModelAdmin):
    list_display = (
        'is_archived', 'name', 'tg_id', 'build_obj', 'salary', 'work_specialization', 'foreman'
    )
    list_display_links = ('name',)
    list_filter = (ArchiveFilter, 'work_specialization', 'build_obj')
    search_fields = ('name', 'tg_id', 'build_obj__name')
    list_editable = ('build_obj', 'work_specialization', 'is_archived')
    fieldsets = (
        (None, {
            'fields': (
                'name', 'tg_id', 'salary', 'work_specialization',
                'foreman', 'build_obj', 'phone_number', 'email',
                'is_archived'
            )
        }),
        ('Advanced options', {
            'classes': ('collapse',),
            'fields': (
                'employment_agreement', 'over_time', 'start_to_work', 'title',
                'payroll_eligible', 'payroll', 'resign_agreement',
                'benefits_eligible', 'birthday', 'issued', 'expiry',
                'address', 'tickets_available'
            ),
        }),
    )
    list_per_page = 20


    def get_actions(self, request):
        """Убираем стандартное удаление, если хочешь запрещать реальное удаление."""
        actions = super().get_actions(request)
        if "delete_selected" in actions:
            del actions["delete_selected"]
        return actions



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


@admin.action(description="Update the tools table")
def update_tools_table_admin(modeladmin, request, queryset):
    for obj in queryset:
        table = ToolsTable(obj)
        table.update_tools_table()

class ToolsSheetAdmin(admin.ModelAdmin):
    fields = ('name', 'sh_url')
    actions = [update_tools_table_admin]



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
@with_month_year_form()
def run_monthly_summary_tasks(modeladmin, request, queryset, month, year):
    all_workers_data = []
    for obj in queryset:
        print(obj.name, month, year)
        try:
            # 1. KPI по пользователям
            table = Sheet1Table(obj)
            workers_data = table.write_summary_workers(month, year)
            logger.info(f"Сводка результатов работников записана для '{obj.name}'")
            all_workers_data.extend(workers_data)
            sleep(5)
            # 2. Сводка по объекту
            table.write_summary(month, year)
            logger.info(f"Сводная таблица за месяц записана для '{obj.name}'")
            sleep(5)

        except Exception as e:
            logger.error(f"Ошибка при обработке '{obj.name}': {e}")
        logger.info(f"Таблицы для объекта {obj.name} записааны, ожидание 60 сек ")


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

class ForemanAndWorkersKPIAdmin(admin.ModelAdmin):
    list_display = ('name', 'sh_url', 'created_at')
    search_fields = ('name',)

@admin.action(description="Update the data of the general statistics tables")
def run_update_general_statistic(modeladmin, request, queryset):
    # Стартуем задачу Celery (один раз, неважно сколько объектов выделено)
    update_general_statistic_task.delay()
    # Выводим сообщение в интерфейсе админки
    messages.success(request, "The task for updating the general statistics has been sent to work.\nThe data will be recorded in a few minutes.")

class GeneralStatisticAdmin(admin.ModelAdmin):
    list_display = ("sheet_name", "sh_url")
    actions = [run_update_general_statistic]


class WorkSpecializationAdmin(admin.ModelAdmin):
    list_display = ["name"]


admin.site.site_header = "TSA"
admin.site.register(WorkSpecialization, WorkSpecializationAdmin)
admin.site.register(GeneralStatisticSheet, GeneralStatisticAdmin)
admin.site.register(MediaProxy, MediaBrowserAdmin)
admin.site.register(ForemanAndWorkersKPISheet, ForemanAndWorkersKPIAdmin)
admin.site.register(Worker, WorkerAdmin)
admin.site.register(BuildBudgetHistory, BuildBudgetHistoryAdmin)
admin.site.register(BuildObject, BuildObjectAdmin)
admin.site.register(Material, MaterialAdmin)
admin.site.register(ToolsSheet, ToolsSheetAdmin)
admin.site.register(Tool, ToolAdmin)
admin.site.register(WorkEntry, WorkEntryAdmin)


