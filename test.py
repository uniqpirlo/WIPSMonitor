# import psutil
import win32evtlog

# [print(id) for id in psutil.pids() if psutil.Process(id).cmdline() == 'WIReportServer.exe']

server = 'localhost'
# server = '10.212.23.173'
logtype = 'Application'
hand = win32evtlog.OpenEventLog(server, logtype)
flags = win32evtlog.EVENTLOG_BACKWARDS_READ | win32evtlog.EVENTLOG_SEQUENTIAL_READ
total = win32evtlog.GetNumberOfEventLogRecords(hand)

monitor_start = '2018-04-05 09:00:00'
quit_flag = False

print(total)

while True:
    events = win32evtlog.ReadEventLog(hand, flags, 0)
    # events = win32evtlog.ReadEventLog(hand, flags, 1)
    # events = win32evtlog.ReadEventLog(hand, flags, 2)
    # events = win32evtlog.ReadEventLog(hand, flags, 0)
    # events = win32evtlog.ReadEventLog(hand, flags, 0)
    if events :
        if quit_flag:
            break
        for event in events:
            # data = event.StringInserts
            # if quit_flag:
            #     break
            if str(event.TimeGenerated) <= monitor_start:
                quit_flag = True
                break
            elif event.EventType == 1:     # 1 means error
                print(event.TimeGenerated, event.EventID, event.SourceName, event.StringInserts)
                # print(dir(event))
            # else:
            #     break
    else:
        break

# print(dir(events[0]))