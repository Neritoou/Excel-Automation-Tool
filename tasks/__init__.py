from typing import TYPE_CHECKING
from tasks.compare_columns import CompareColumnsTask

if TYPE_CHECKING:
    from core.base_task import BaseTask

ALL_TASKS: list[type["BaseTask"]] = [CompareColumnsTask]