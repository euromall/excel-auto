"""Progress reporting for long work (reading big workbooks, writing files).

Workers call step()/update(); the window shows them in a progress popup. Without a
window (CLI, tests) a Progress() does nothing. Cancel is honoured only in steps
marked cancellable, so a file is never left half written.
"""


class Cancelled(Exception):
    pass


class Progress:
    def __init__(self, on_change=None):
        self.on_change = on_change          # fn(step, steps, text, done, total, cancellable)
        self.cancel_requested = False
        self.step_no, self.steps, self.text, self.cancellable = 0, 0, '', True

    def step(self, no: int, steps: int, text: str, cancellable: bool = True):
        self.check()
        self.step_no, self.steps, self.text, self.cancellable = no, steps, text, cancellable
        self._emit(0, 0)

    def update(self, done: int, total: int):
        """done/total within the current step (total 0 = unknown)."""
        self.check()
        self._emit(done, total)

    def check(self):
        if self.cancel_requested and self.cancellable:
            raise Cancelled()

    def _emit(self, done, total):
        if self.on_change:
            self.on_change(self.step_no, self.steps, self.text, done, total, self.cancellable)


def tick(progress, i: int, total, every: int = 200):
    """Report row i of total every `every` rows (cheap enough for row loops)."""
    if progress is not None and i % every == 0:
        progress.update(i, total or 0)
