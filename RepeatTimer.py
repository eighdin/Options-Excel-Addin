from threading import Timer

class RepeatTimer:
    def __init__(self, interval, function, *args, **kwargs):
        self._timer = None
        self.interval = interval
        self.function = function
        self.args = args
        self.kwargs = kwargs
        self.is_running = False
        self.start()

    def _run(self):
        self.is_running = False
        self.function(*self.args, **self.kwargs)
        self.start()

    def start(self):
        if not self.is_running:
            self._timer = Timer(self.interval*60, self._run)
            self._timer.start()
            self.is_running = True

    def stop(self):
        if self._timer:
            self._timer.cancel()
        self.is_running = False
