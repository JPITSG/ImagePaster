/* Run elevated on Windows; uses an isolated task name and removes it on exit.
   Build: mingw-gcc tests/windows_startup_task_test.c -o startup-test.exe
          -lole32 -loleaut32 -ladvapi32 -luuid -ltaskschd */
#define IMAGEPASTER_STARTUP_TASK_PREFIX L"ImagePaster-StartupTest-"
#include "../startup_task.c"

#define CHECK(condition) do { if (!(condition)) { \
    fprintf(stderr, "Failed line %d: %s\n", __LINE__, #condition); \
    failed = 1; goto done; } } while (0)

int main(void)
{
    int failed = 0;
    BOOL present = FALSE, enabled = FALSE;
    StartupTaskContext context = {0};
    IRegisteredTask *task = NULL;
    BOOL created = FALSE;
    HRESULT initialized = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    CHECK(SUCCEEDED(initialized));
    CHECK(ReadStartupTask(&present, &enabled) == ERROR_SUCCESS && !present && !enabled);
    CHECK(WriteStartupTask(TRUE) == ERROR_SUCCESS);
    created = TRUE;
    CHECK(ReadStartupTask(&present, &enabled) == ERROR_SUCCESS && present && enabled);
    CHECK(OpenStartupTask(&context) == ERROR_SUCCESS);
    CHECK(SUCCEEDED(ITaskFolder_GetTask(context.folder, context.name, &task)));
    CHECK(SUCCEEDED(IRegisteredTask_put_Enabled(task, VARIANT_FALSE)));
    CHECK(ReadStartupTask(&present, &enabled) == ERROR_SUCCESS && present && !enabled);
    CHECK(WriteStartupTask(TRUE) == ERROR_SUCCESS);
    CHECK(ReadStartupTask(&present, &enabled) == ERROR_SUCCESS && present && enabled);
    /* Inspect another executable's ownership without modifying the actual task. */
    context.path[0] = context.path[0] == L'Z' ? L'Y' : L'Z';
    CHECK(SUCCEEDED(StartupTaskMatches(task, &context, &enabled)) && !enabled);
    CHECK(WriteStartupTask(FALSE) == ERROR_SUCCESS);
    CHECK(ReadStartupTask(&present, &enabled) == ERROR_SUCCESS && !present && !enabled);
    CHECK(WriteStartupTask(FALSE) == ERROR_SUCCESS);
done:
    if (created && WriteStartupTask(FALSE) != ERROR_SUCCESS) failed = 1;
    if (task) IRegisteredTask_Release(task);
    CloseStartupTask(&context);
    if (SUCCEEDED(initialized)) CoUninitialize();
    if (!failed) puts("PASS: task creation, identity, path, elevation, disable, repair, ownership and removal");
    return failed;
}
