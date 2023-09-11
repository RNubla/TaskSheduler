# import win32com.client
# from nicegui import ui


# class TaskScheduler:
#     def __init__(self) -> None:
#         self.scheduler = win32com.client.Dispatch("Schedule.Service")
#         self.scheduler.Connect()
#         self.folders = None
#         self.TASK_ENUM_HIDDEN = 1
#         self.TASK_STATE = {
#             0: "Unknown",
#             1: "Disabled",
#             2: "Queued",
#             3: "Ready",
#             4: "Running",
#         }
#         self.lastRunResultMsg = {
#             "2": "(0x2)",
#             "0": "The operation completed successfully (0x0)",
#             "1": "(0x1)",
#             "267011": "The task has not yet run. (0x41303)",
#         }
#         self.jobs = []
#         self.grid = ui.aggrid(
#             {
#                 "defaultColDef": {"flex": 1},
#                 "columnDefs": [
#                     {"headerName": "Path", "field": "path"},
#                     {"headerName": "Name", "field": "name"},
#                     {"headerName": "Status", "field": "state"},
#                     {"headerName": "Last Run", "field": "lastRun"},
#                     {"headerName": "Next Run", "field": "nextRun"},
#                     {"headerName": "Last Result", "field": "lastResult"},
#                 ],
#                 "rowData": self.jobs,
#             }
#         ).classes("h-screen")

#     def clearTable(self):
#         print("clearTable")
#         # self.grid.clear()
#         self.jobs = []
#         self.grid.options["rowData"] = []
#         self.grid.update()

#     def getLastRunMsg(self, msg: str) -> str:
#         return self.lastRunResultMsg[msg] if msg in self.lastRunResultMsg else msg

#     def fetchAllJobs(self) -> None:
#         self.clearTable()
#         self.folders = [self.scheduler.GetFolder("\\")]
#         while self.folders:
#             folder = self.folders.pop(0)
#             self.folders += list(folder.GetFolders(0))
#             tasks = list(folder.GetTasks(self.TASK_ENUM_HIDDEN))
#             for task in tasks:
#                 self.jobs.append(
#                     {
#                         "path": task.Path,
#                         "name": task.Name,
#                         "state": self.TASK_STATE[task.State],
#                         "lastRun": str(task.LastRunTime),
#                         "nextRun": str(task.NextRunTime),
#                         "lastResult": self.getLastRunMsg(str(task.LastTaskResult)),
#                     }
#                 )
#         self.grid.options["rowData"] = self.jobs
#         self.grid.update()

#     def fetchAllJobsExcludeFolder(self, folderName: str) -> None:
#         self.clearTable()
#         self.folders = [self.scheduler.GetFolder("\\")]
#         while self.folders:
#             folder = self.folders.pop(0)
#             self.folders += list(folder.GetFolders(0))
#             tasks = list(folder.GetTasks(self.TASK_ENUM_HIDDEN))
#             for task in tasks:
#                 if folderName not in task.Path:
#                     self.jobs.append(
#                         {
#                             "path": task.Path,
#                             "name": task.Name,
#                             "state": self.TASK_STATE[task.State],
#                             "lastRun": str(task.LastRunTime),
#                             "nextRun": str(task.NextRunTime),
#                             "lastResult": str(task.LastTaskResult),
#                         }
#                     )
#         self.grid.options["rowData"] = self.jobs
#         self.grid.update()


# if __name__ in {"__main__", "__mp_main__"}:
#     ui.label("Hello NiceGUI!")
#     ui.button(
#         # "Fetch Job List", on_click=lambda: app.fetchAllJobsExcludeFolder("Microsoft")
#         "Fetch Job List",
#         on_click=lambda: app.fetchAllJobs(),
#     )
#     # ui.button("Clear Table", on_click=lambda: app.clearTable())
#     app = TaskScheduler()
#     ui.run()

from datetime import datetime, timedelta
import multiprocessing
import re
import uuid
import win32com.client
import win32api
from fastapi import FastAPI
import uvicorn
from fastapi.middleware.cors import CORSMiddleware


class TaskSchedulerService:
    def __init__(self) -> None:
        self.task_scheduler = win32com.client.Dispatch("Schedule.Service")
        self.task_scheduler.Connect()
        self.root_folder = self.task_scheduler.GetFolder("\\")

        self.task_list = []

    def read_tasks_in_folder(self, folder):
        for task in folder.GetTasks(0):
            self.append_tasks_list(task=task)

        # Recursively list task in subfolders
        for subfolder in folder.GetFolders(0):
            self.read_tasks_in_folder(subfolder)

    def append_tasks_list(self, task):
        if "Microsoft" not in task.Path and "Optimize" not in task.Path:
            self.task_list.append(
                {
                    "id": uuid.uuid4(),
                    "path": task.Path,
                    "name": task.Name,
                    "state": self.get_task_state_string(task.State),
                    "lastRunTime": self.format_time(task.LastRunTime),
                    "nextRunTime": self.format_time(task.NextRunTime),
                    "lastTaskResult": self.get_error_message(task.LastTaskResult),
                    "triggerDetails": self.get_task_trigger(task=task),
                }
            )

    def get_error_message(self, erro_code: int) -> str:
        """Returns the message representation of the error_code"""
        try:
            message = win32api.FormatMessage(erro_code).strip()
            return message
        except Exception as e:
            return f"Error Code {erro_code}. Reason: {str(e)}"

    def get_task_list(self):
        return self.task_list

    def get_task_state_string(self, task_state):
        """Return a string representation of the task state."""
        state_map = {0: "Unknown", 1: "Disabled", 2: "Queued", 3: "Ready", 4: "Running"}

        return state_map.get(task_state, "Unknown")

    def get_trigger_details(self, trigger):
        """Return trigger details in dictionary or object format"""
        trigger_type = trigger.Type
        details = {
            "type": None,
            "startBoundary": None,
            "endBoundary": "indefinitely",
            "duration": None,
            "repetition": None,
            "enabled": False,
        }

        match trigger_type:
            case 1:  # TASK_TRIGGER_TIME
                details["startBoundary"] = trigger.StartBoundary
                details["enabled"] = trigger.Enabled
                details["type"] = "once"
            case 2:  # TASK_TRIGGER_DAILY
                details["type"] = "daily"
                details["startBoundary"] = trigger.StartBoundary
                details["enabled"] = trigger.Enabled
                details["duration"] = f"every {trigger.DaysInterval} day(s)"
                if trigger.Repetition.Duration:
                    details[
                        "repetition"
                    ] = f"every {self.parse_iso_duration(trigger.Repetition.Interval)} for {self.parse_iso_duration(trigger.Repetition.Duration)}"
            case 3:  # TASK_TRIGGER_WEEKLY
                details["type"] = "weekly"
                details["startBoundary"] = trigger.StartBoundary
                details["enabled"] = trigger.Enabled
                details[
                    "duration"
                ] = f"every {trigger.WeeksInterval} week(s) on { self.decode_days_of_week(trigger.DaysOfWeek)}"
                if trigger.Repetition.Duration:
                    details[
                        "repetition"
                    ] = f"every {self.parse_iso_duration(trigger.Repetition.Interval)} for {self.parse_iso_duration(trigger.Repetition.Duration)}"

            case 4:  # TASK_TRIGGER_MONTHLY
                details["type"] = "monthly"
                details["startBoundary"] = trigger.StartBoundary
                details["enabled"] = trigger.Enabled
                details[
                    "duration"
                ] = f"every {trigger.MonthsOfYear} month(s) on day {trigger.DaysOfMonth}"
                if trigger.Repetition.Duration:
                    details[
                        "repetition"
                    ] = f"every {self.parse_iso_duration(trigger.Repetition.Interval)} for {self.parse_iso_duration(trigger.Repetition.Duration)}"

            case 5:  # TASK_TRIGGER_IDLE
                details["type"] = "monthlyDow"
                details["startBoundary"] = trigger.StartBoundary
                details["enabled"] = trigger.Enabled
                details[
                    "duration"
                ] = f"every {trigger.MonthsOfYear} month(s) on {self.decode_days_of_week(trigger.DaysOfWeek)} of week {trigger.WeeksOfMonth}"
                if trigger.Repetition.Duration:
                    details[
                        "repetition"
                    ] = f"every {self.parse_iso_duration(trigger.Repetition.Interval)} for {self.parse_iso_duration(trigger.Repetition.Duration)}"

            case 6:  # TASK_TRIGGER_IDLE
                details["type"] = "idle"
                details["enabled"] = trigger.Enabled
                # No specific start boundary for Idle triggers, as they start when the system goes idle.

            case 7:  # TASK_TRIGGER_LOGON
                details["type"] = "registration"
                details["startBoundary"] = trigger.StartBoundary
                details["enabled"] = trigger.Enabled

            case 8:  # TASK_TRIGGER_LOGON
                details["type"] = "boot"
                details["enabled"] = trigger.Enabled
                # No specific start boundary for Boot triggers, as they start when the system boots.
            case 9:  # TASK_TRIGGER_LOGON
                details["type"] = "logon"
                details["startBoundary"] = trigger.StartBoundary
                details["enabled"] = trigger.Enabled

            case 11:  # TASK_TRIGGER_SESSION_STATE_CHANGE
                details["type"] = "session state change"
                details["startBoundary"] = trigger.StartBoundary
                details["enabled"] = trigger.Enabled
            case _:
                details["type"] = "unknown"

        if trigger.EndBoundary:
            details["endBoundary"] = trigger.EndBoundary
        else:
            details["endBoundary"] = "indefinitely"

        return details

    def get_task_trigger(self, task):
        """Return trigger details for a given task"""
        triggers = []
        for trigger in task.Definition.Triggers:
            trigger_details = self.get_trigger_details(trigger)
            triggers.append(trigger_details)
        return triggers

    def format_time(self, task_time):
        return datetime.strptime(str(task_time)[:19], "%Y-%m-%d %H:%M:%S").strftime(
            "%Y-%m-%d %I:%M:%S %p"
        )

    def decode_days_of_week(self, days_value):
        days = [
            "Sunday",
            "Monday",
            "Tuesday",
            "Wednesday",
            "Thursday",
            "Friday",
            "Saturday",
        ]
        scheduled_days = []

        for i, day in enumerate(days):
            if days_value & (1 << i):
                scheduled_days.append(day)

        return ", ".join(scheduled_days)

    def parse_iso_duration(self, duration_str):
        """Convert ISO 8601 duration format (like PT10M) to a timedelta."""
        pattern = re.compile(r"(?P<hours>\d+H)?(?P<minutes>\d+M)?(?P<seconds>\d+S)?")
        parts = pattern.search(duration_str)

        if not parts:
            return timedelta()

        parts = parts.groupdict()
        hours = int(parts["hours"][:-1]) if parts["hours"] else 0
        minutes = int(parts["minutes"][:-1]) if parts["minutes"] else 0
        seconds = int(parts["seconds"][:-1]) if parts["seconds"] else 0

        return timedelta(hours=hours, minutes=minutes, seconds=seconds)


app = FastAPI()

origins = [
    "http://localhost.tiangolo.com",
    "https://localhost.tiangolo.com",
    "http://localhost",
    "http://localhost:3000",
]
app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/")
async def root():
    return {"message": "Hello World"}


@app.get("/api/TaskScheduler")
async def TaskScheduler():
    ts = TaskSchedulerService()
    ts.read_tasks_in_folder(ts.root_folder)
    return ts.task_list


if __name__ == "__main__":
    multiprocessing.freeze_support()
    uvicorn.run("main:app", host="0.0.0.0", port=8080, reload=True)
