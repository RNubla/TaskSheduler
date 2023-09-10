import win32evtlog


def get_task_scheduler_logs(
    log_type="Operational", log_source="Microsoft-Windows-TaskScheduler"
):
    server = "localhost"  # Local machine
    log_handle = win32evtlog.OpenEventLog(server, log_type)

    flags = win32evtlog.EVENTLOG_BACKWARDS_READ | win32evtlog.EVENTLOG_SEQUENTIAL_READ
    events = win32evtlog.ReadEventLog(log_handle, flags, 0)

    results = []
    for event in events:
        if event.SourceName == log_source:
            data = {
                "EventID": event.EventID,
                "TimeGenerated": event.TimeGenerated,
                "StringInserts": event.StringInserts,
            }
            results.append(data)

    win32evtlog.CloseEventLog(log_handle)
    return results


if __name__ == "__main__":
    logs = get_task_scheduler_logs()
    for log in logs:
        print(log)
