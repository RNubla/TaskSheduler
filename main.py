import win32com.client
from nicegui import ui


class TaskScheduler:
    def __init__(self) -> None:
        self.scheduler = win32com.client.Dispatch("Schedule.Service")
        self.scheduler.Connect()
        self.folders = None
        self.TASK_ENUM_HIDDEN = 1
        self.TASK_STATE = {
            0: "Unknown",
            1: "Disabled",
            2: "Queued",
            3: "Ready",
            4: "Running",
        }
        self.jobs = []
        self.grid = ui.aggrid(
            {
                "defaultColDef": {"flex": 1},
                "columnDefs": [
                    {"headerName": "Path", "field": "path"},
                    {"headerName": "Name", "field": "name"},
                    {"headerName": "Status", "field": "state"},
                    {"headerName": "Last Run", "field": "lastRun"},
                    {"headerName": "Next Run", "field": "nextRun"},
                    {"headerName": "Last Result", "field": "lastResult"},
                ],
                "rowData": self.jobs,
            }
        ).classes("h-screen")

    # def connect(self):

    def fetchAllJobs(self) -> None:
        self.folders = [self.scheduler.GetFolder("\\")]
        while self.folders:
            folder = self.folders.pop(0)
            self.folders += list(folder.GetFolders(0))
            tasks = list(folder.GetTasks(self.TASK_ENUM_HIDDEN))
            for task in tasks:
                self.jobs.append(
                    {
                        "path": task.Path,
                        "name": task.Name,
                        "state": self.TASK_STATE[task.State],
                        "lastRun": task.LastRunTime,
                        "nextRun": task.NextRunTime,
                        "lastResult": task.LastTaskResult,
                    }
                )
        self.grid.update()

    def fetchAllJobsExcludeFolder(self, folderName: str) -> None:
        # self.jobs = []
        # self.grid.clear()
        self.folders = [self.scheduler.GetFolder("\\")]
        while self.folders:
            folder = self.folders.pop(0)
            self.folders += list(folder.GetFolders(0))
            tasks = list(folder.GetTasks(self.TASK_ENUM_HIDDEN))
            for task in tasks:
                if folderName not in task.Path:
                    self.jobs.append(
                        {
                            "path": task.Path,
                            "name": task.Name,
                            "state": self.TASK_STATE[task.State],
                            "lastRun": str(task.LastRunTime),
                            "nextRun": str(task.NextRunTime),
                            "lastResult": str(task.LastTaskResult),
                        }
                    )
        # print("grid update", self.jobs)
        self.grid.update()
        # self.grid.clear()
        # self.grid.clear()

    def clearTable(self):
        # self.grid.options["rowData"][0]
        print(self.grid.options["rowData"])
        # self.grid.options['rowData'].remove
        # self.grid.call_api_method("setRowData", [None])


if __name__ in {"__main__", "__mp_main__"}:
    ui.label("Hello NiceGUI!")
    ui.button("Update", on_click=lambda: app.fetchAllJobsExcludeFolder("Microsoft"))
    ui.button("Clear Table", on_click=lambda: app.clearTable)
    app = TaskScheduler()
    ui.run()
