import sys
import win32gui, win32process, _thread, queue, time
import configparser
import smtplib
import email.message, email.utils

def getwindowinfo() -> (int, str, int, int):
    "Returns (hwnd, text, threadId, processId)"
    hwnd = win32gui.GetForegroundWindow()
    text = win32gui.GetWindowText(hwnd)
    threadid, procid = win32process.GetWindowThreadProcessId(hwnd)
    return hwnd, text, threadid, procid

def input_thread(log_queue:queue.Queue, command_queue:queue.Queue, stdout_mutex:_thread.LockType):
    global exited_threads
    while not exited_threads2:
        line = sys.stdin.readline().strip()
        command_queue.put(line)
        if line == 'exit':
            exited_threads += 1
            break

def log_thread(log_queue:queue.Queue, log_queue2:queue.Queue, config:configparser.ConfigParser, log:dict, log2:dict, wait_time:int):
    global exited_threads
    file = open(log['filename'], 'a', buffering=1)
    while not exited_threads2:
        while not log_queue.empty():
            m = log_queue.get()
            try: print(m, file=file)
            except UnicodeEncodeError: print(repr(m), file=file)
            log_queue2.put(m)
        if time.time() >= log['endtime']:
            file.close()
            log2.clear()
            log2['starttime'] = time.time()
            log2['endtime'] = log2['starttime'] + parse_length(
                config.getfloat('logging', 'new_log_time', fallback=0),
                config.get('logging', 'when_new', fallback='never'))
            log2.update({'startdate': time.ctime(log2['starttime']), 'enddate': time.ctime(log2['endtime'])})
            log2['filename'] = config.get('logging', 'filename', fallback='screenlog.log') % log2
            log2['filename'] = log2['filename'].replace(':', '_').replace('/', '_').replace('\\', '_')
            file = open(log2['filename'], 'a', buffering=1)
    file.close()
    exited_threads += 1

def doprint(mutex:_thread.LockType, *vals, sep=' '):
    mutex.acquire()
    print('\r' + vals[0], *vals[1:], end='\n> ', sep=sep)
    mutex.release()

exited_threads = None
exited_threads2 = None

def parse_length(new_log_time, when_new):
    if when_new == 'never':
        return 0
    elif when_new == 'timeloop_minutes':
        return new_log_time * 60
    elif when_new == 'timeloop_hours':
        return new_log_time * 60 * 60
    elif when_new == 'timeloop_days':
        return new_log_time * 60 * 60 * 24
    else:
        return new_log_time

def log_message(queue:queue.Queue, message):
    queue.put('%s -- %s' % (time.ctime(), message))

def detect_thread(log_queue:queue.Queue, config:configparser.ConfigParser, wait_time:int):
    global exited_threads
    option = config.get('tracking', 'log_level', fallback='window')
    prevvalue = None
    while not exited_threads2:
        time.sleep(wait_time)
        window = getwindowinfo()
        if option == 'window':
            value = window[0]
        elif option == 'title':
            value = window[1]
        elif option == 'thread':
            value = window[2]
        elif option == 'process':
            value = window[3]
        else:
            value = None
        if value != prevvalue:
            label = window[1].strip()
            if label:
                prevvalue = value
                log_message(log_queue, label)

def run(config:configparser.ConfigParser, configfile:str=None):
    global exited_threads, exited_threads2
    exited_threads = 0
    exited_threads2 = 0
    wait_time = config.getint('tracking', 'poll_time', fallback=100)/1000.
    log_queue = queue.Queue()
    log_queue2 = queue.Queue()
    command_queue = queue.Queue()
    stdout_mutex = _thread.allocate_lock()
    _thread.start_new_thread(input_thread, (log_queue, command_queue, stdout_mutex))
    log = {
        'starttime': time.time(),
    }
    log['endtime'] = log['starttime'] + parse_length(
        config.getfloat('logging', 'new_log_time', fallback=0),
        config.get('logging', 'when_new', fallback='never'))
    log.update({'startdate': time.ctime(log['starttime']), 'enddate': time.ctime(log['endtime'])})
    log['filename'] = config.get('logging', 'filename', fallback='screenlog.log') % log
    log['filename'] = log['filename'].replace(':', '_').replace('/', '_').replace('\\', '_')
    log2 = {}
    _thread.start_new_thread(log_thread, (log_queue, log_queue2, config, log, log2, wait_time))
    
    try:
        _thread.start_new_thread(detect_thread, (log_queue, config, wait_time))
        log_message(log_queue, 'Screen Log started...')
        while not exited_threads:
            time.sleep(wait_time)
            while not command_queue.empty():
                command = command_queue.get()
                if command.startswith('log'):
                    arg = command.split(' ', 1)[1]
                    log_message(log_queue, arg)
                elif command == 'reloadconf':
                    if configfile is None:
                        doprint(stdout_mutex, 'Sorry, the config file was not provided')
                    else:
                        config.clear()
                        config.read(configfile)
                        doprint(stdout_mutex, 'Successfully reloaded config...')
                # elif command == 'exit_no_email':
                #     config['email']['do_email'] = 'no'
                #     break
                elif command == 'forcemail':
                    send_email(config, log)
                else:
                    doprint(stdout_mutex, 'Invalid command "%s"' % command)
            while not log_queue2.empty():
                doprint(stdout_mutex, log_queue2.get())
            if log2 != {}:
                send_email(config, log)
                log.update(log2)
                log2.clear()
        log_message(log_queue, 'Screen Log exited...')
        exited_threads += 1
        time.sleep(wait_time)
        exited_threads2 = exited_threads
        log['endtime'] = time.time()
        log['enddate'] = time.ctime()
        send_email(config, log)
    except:
        log_message(log_queue, 'Screen Log exited...')
        exited_threads += 1
        time.sleep(wait_time)
        exited_threads2 = exited_threads
        # log['endtime'] = time.time()
        # log['enddate'] = time.ctime()
        # send_email(config, log)
        raise

def send_email(config:configparser.ConfigParser, log:dict):
    if not config.getboolean('email', 'do_email', fallback=False):
        return

    message = email.message.EmailMessage()
    message['From'] = config.get('email', 'from')
    message['To'] = config.get('email', 'to')
    if config.has_option('email', 'cc'): message['Cc'] = config.get('email', 'cc')
    if config.has_option('email', 'bcc'): message['Bcc'] = config.get('email', 'bcc')
    message['Date'] = email.utils.formatdate(localtime=True)
    message['Subject'] = config.get('email', 'subject') % log
    textmessage = email.message.Message()
    textmessage.set_payload(config.get('email', 'body') % log)
    message.make_mixed()
    message.attach(textmessage)
    message.add_attachment(open(log['filename'], 'rb').read(), filename=log['filename'], maintype='text', subtype='plain')

    smtp = smtplib.SMTP(config.get('email', 'smtp_server'), config.getint('email', 'smtp_port', fallback=25))
    smtp.ehlo()
    smtp.starttls()
    smtp.login(config.get('email', 'username'), config.get('email', 'password'))
    smtp.send_message(message)
    smtp.quit()

def main():
    p = configparser.ConfigParser(interpolation=None)
    fname = p.read((len(sys.argv) > 1 and sys.argv[1]) or 'screenlog.ini')
    run(p, fname)

if __name__ == '__main__':
    # from configparser import ConfigParser
    # p = ConfigParser(interpolation=None)
    # p.read('screenlog.ini')
    # send_email(p, {'startdate':0, 'enddate':10000, 'filename': 'screenlog today - tomorrow.log'})
    main()