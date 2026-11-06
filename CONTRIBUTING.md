# Contributing

    uv sync
    uv run pytest
    uv run ruff check src tests

The tests run on macOS, Windows and Linux; none of them needs Outlook. The macOS
backend runs against a fake script runner (`tests/mac/conftest.py`), the Windows
backend against a fake Outlook object model (`tests/windows/fakecom.py`). When you
find that real Outlook behaves differently from the fake, change the fake first
so a test shows the difference, then fix the backend.

Rules the code keeps, and that the tests check:

- No tool runs caller-supplied script code. Caller text reaches Outlook only as
  data (`argv` of the fixed AppleScripts, property values over COM).
- Reading changes nothing: read AppleScripts may use no changing command, and on
  Windows each function may use only the object-model calls listed for it in
  `tests/windows/test_windows_safety.py`. A new changing call needs an entry
  there, which makes it visible in review.
- Nothing deletes permanently. On Windows, `Delete()` is not called anywhere.
- Anything that sends to other people is a separate, explicit step.

Please test changes to the AppleScripts or the COM code on a real Outlook as
well, and say in the pull request which Outlook version you used.
