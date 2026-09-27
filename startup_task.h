#ifndef IMAGEPASTER_STARTUP_TASK_H
#define IMAGEPASTER_STARTUP_TASK_H

#include <windows.h>

/* COM must be initialized on the calling thread. The task belongs to the
   current user, uses their interactive token, and launches this executable. */
LONG ReadStartupTask(BOOL *present, BOOL *enabled);
LONG WriteStartupTask(BOOL enable);

#endif
