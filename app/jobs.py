"""In-memory job store for backgrounded parses.

WHY THIS EXISTS. A real PHPP parse takes ~46 seconds — measured in production on
a 9.7 MB easyPH workbook. Heroku's router caps time-to-first-byte at 30 seconds,
which is not configurable. So a synchronous ``POST /parse`` can never succeed for
a real workbook: the router hangs up at 30s while the work completes behind it,
and the caller sees a 503 with an HTML body it did not ask for. The production
logs show exactly that pair — ``H12`` at 30s, ``"POST /parse" 200 OK`` sixteen
seconds later.

A larger dyno does not close that gap. The service was moved Standard-1X ->
Standard-2X while diagnosing this and the elapsed time did not move; the cost is
openpyxl reading a 9.7-32.8 MB workbook, not CPU contention.

So ``/parse`` accepts the work and returns 202 with a job id, and the caller
polls ``/parse/{job_id}``.

STATE IS DELIBERATELY IN-PROCESS, and that is a decision rather than an
oversight:

  * this service runs a SINGLE uvicorn worker, so there is no cross-process
    sharing problem that a shared store would solve;
  * jobs are short-lived, and the only consumer is a shadow path in ps-rails
    which feeds an offline comparison — nothing user-facing waits on it. Losing
    in-flight jobs when Heroku cycles the dyno (roughly daily) costs a retry,
    not correctness.

If the parse ever becomes user-facing, this needs durable state and the caller
wants a callback rather than a poll. That design already exists next door, for
the PHPP export.
"""

from __future__ import annotations

import threading
import time
import uuid
from dataclasses import dataclass, field
from typing import Any, Literal

Status = Literal["pending", "done", "failed"]

# How long a finished job stays readable. Long enough that a poller which backs
# off, or a dyno under load, still finds its result; short enough that the dict
# cannot grow without bound on a long-lived process.
TTL_SECONDS = 900


@dataclass
class Job:
    status: Status = "pending"
    result: dict[str, Any] | None = None
    detail: str | None = None
    created_at: float = field(default_factory=time.monotonic)
    finished_at: float | None = None


# The parse runs in a worker thread and writes here; request handlers read here.
# Different threads, so the lock is load-bearing, not decorative.
_jobs: dict[str, Job] = {}
_lock = threading.Lock()


def create() -> str:
    """Register a new pending job and return its id."""
    job_id = uuid.uuid4().hex
    with _lock:
        _evict_expired_locked()
        _jobs[job_id] = Job()
    return job_id


def get(job_id: str) -> Job | None:
    with _lock:
        return _jobs.get(job_id)


def finish(job_id: str, result: dict[str, Any]) -> None:
    with _lock:
        job = _jobs.get(job_id)
        if job is None:  # evicted, or the process restarted under us
            return
        job.status = "done"
        job.result = result
        job.finished_at = time.monotonic()


def fail(job_id: str, detail: str) -> None:
    with _lock:
        job = _jobs.get(job_id)
        if job is None:
            return
        job.status = "failed"
        job.detail = detail
        job.finished_at = time.monotonic()


def _evict_expired_locked() -> None:
    """Drop finished jobs past their TTL. Caller must hold the lock.

    Only FINISHED jobs are evicted. A pending job is never dropped on age: a
    slow parse that outlived the TTL is precisely the case this service exists
    to support, and evicting it would hand the poller a 404 that looks like a
    lost job rather than a slow one.
    """
    cutoff = time.monotonic() - TTL_SECONDS
    expired = [
        job_id
        for job_id, job in _jobs.items()
        if job.finished_at is not None and job.finished_at < cutoff
    ]
    for job_id in expired:
        del _jobs[job_id]


def _reset_for_tests() -> None:
    with _lock:
        _jobs.clear()
