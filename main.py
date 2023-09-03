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
        self.lastRunResultMsg = {
            "2": "(0x2)",
            "0": "The operation completed successfully (0x0)",
            "1": "(0x1)",
            "267011": "The task has not yet run. (0x41303)",
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

    def clearTable(self):
        print("clearTable")
        # self.grid.clear()
        self.jobs = []
        self.grid.options["rowData"] = []
        self.grid.update()

    def getLastRunMsg(self, msg: str) -> str:
        return self.lastRunResultMsg[msg] if msg in self.lastRunResultMsg else msg

    def fetchAllJobs(self) -> None:
        self.clearTable()
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
                        "lastRun": str(task.LastRunTime),
                        "nextRun": str(task.NextRunTime),
                        "lastResult": self.getLastRunMsg(str(task.LastTaskResult)),
                    }
                )
        self.grid.options["rowData"] = self.jobs
        self.grid.update()

    def fetchAllJobsExcludeFolder(self, folderName: str) -> None:
        self.clearTable()
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
        self.grid.options["rowData"] = self.jobs
        self.grid.update()


if __name__ in {"__main__", "__mp_main__"}:
    ui.label("Hello NiceGUI!")
    ui.button(
        # "Fetch Job List", on_click=lambda: app.fetchAllJobsExcludeFolder("Microsoft")
        "Fetch Job List",
        on_click=lambda: app.fetchAllJobs(),
    )
    # ui.button("Clear Table", on_click=lambda: app.clearTable())
    app = TaskScheduler()
    ui.run()
