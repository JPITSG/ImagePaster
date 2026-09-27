# Input over elevated windows — 1.0.46

On 2026-09-27, ImagePaster 1.0.45 on two Windows PCs received Print Screen over
ordinary windows but failed over an elevated SystrayLauncher on one PC and
MonitorBuddy's elevated shared-window proxies on either PC. The local
SystrayLauncher failure also occurred with MonitorBuddy stopped. Foreground
activation succeeded; this was not a missing-focus problem.

ImagePaster was running at medium integrity. SystrayLauncher on the failing PC
and MonitorBuddy's proxies ran at high integrity. Controlled input tests found:

| Listener | Foreground | Result |
| --- | --- | --- |
| ImagePaster, medium | Taskbar / ordinary Firefox | Capture opened |
| ImagePaster, medium | Elevated SystrayLauncher | No capture |
| Separate F24 hook + hotkey probe, medium | Elevated SystrayLauncher | 0 hook events, 0 hotkeys |
| Separate F24 hook + hotkey probe, high | Elevated SystrayLauncher | 2 hook events, 1 hotkey |
| ImagePaster, high | Taskbar / Firefox / SystrayLauncher, both PCs | Capture opened |
| ImagePaster, high | Actual shared Firefox proxy, either PC | Capture opened on the receiving PC |

The separate F24 probe distinguishes a general input privilege boundary from
a Print Screen-specific conflict. Changing the injecting process's privileges
did not cure the medium-integrity ImagePaster failure; elevating the listener did.
Tests used marked synthetic input, verified the foreground HWND and local capture
process, and dismissed overlays without taking a capture. These tests validate
input delivery on the selected desktop, not a physical keyboard's complete
cross-PC route. MonitorBuddy's existing pointer-based system-key routing is
unchanged.

Version 1.0.46 requests administrator privileges for the whole input handler,
covering hooks, registered hotkeys and paste injection. Its optional sign-in
startup uses Task Scheduler's per-user interactive token with highest available
privileges, replacing the old Run entry. It does not store credentials or use
SYSTEM. An enabled legacy entry for this copy migrates only after successful
task creation; disabled settings and other copies' entries are preserved.

The host regression suite covers migration and failure handling. The native
`tests/windows_startup_task_test.c` passed on both PCs from paths containing
spaces and `&`, checking real task creation, account identity normalization,
executable ownership, required elevation, disabled state, re-enabling and
removal. Its isolated test task is deleted on exit.

Windows API references: [application manifests](https://learn.microsoft.com/en-us/windows/win32/sbscs/application-manifests),
[Task Scheduler security contexts](https://learn.microsoft.com/en-us/windows/win32/taskschd/security-contexts-for-running-tasks),
and [SendInput integrity restrictions](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-sendinput).
