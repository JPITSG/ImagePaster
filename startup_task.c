#define UNICODE
#define _UNICODE
#define _WIN32_WINNT 0x0600
#define COBJMACROS
#include "startup_task.h"
#include <objbase.h>
#include <oleauto.h>
#include <sddl.h>
#include <taskschd.h>
#include <stdio.h>

#ifndef IMAGEPASTER_STARTUP_TASK_PREFIX
#define IMAGEPASTER_STARTUP_TASK_PREFIX L"ImagePaster-Startup-"
#endif

typedef struct StartupTaskContext {
    ITaskService *service;
    ITaskFolder *folder;
    BSTR name, sid;
    wchar_t path[MAX_PATH];
} StartupTaskContext;

static LONG TaskError(HRESULT result)
{
    return SUCCEEDED(result) ? ERROR_SUCCESS :
        HRESULT_FACILITY(result) == FACILITY_WIN32 ? HRESULT_CODE(result) : result;
}

static void CloseStartupTask(StartupTaskContext *context)
{
    if (context->folder) ITaskFolder_Release(context->folder);
    if (context->service) ITaskService_Release(context->service);
    SysFreeString(context->name);
    SysFreeString(context->sid);
}

static LONG OpenStartupTask(StartupTaskContext *context)
{
    HANDLE token = NULL;
    BYTE user[1024];
    DWORD needed = 0;
    wchar_t *sid = NULL;
    wchar_t name[256];
    VARIANT empty;
    BSTR root;
    HRESULT result;
    DWORD length = GetModuleFileNameW(NULL, context->path, MAX_PATH);
    if (!length || length >= MAX_PATH) return ERROR_BAD_PATHNAME;
    if (!OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, &token)) return GetLastError();
    BOOL ok = GetTokenInformation(token, TokenUser, user, sizeof(user), &needed) &&
        ConvertSidToStringSidW(((TOKEN_USER *)user)->User.Sid, &sid);
    LONG error = ok ? ERROR_SUCCESS : GetLastError();
    CloseHandle(token);
    if (!ok) return error;
    int count = swprintf(name, sizeof(name) / sizeof(*name),
                        IMAGEPASTER_STARTUP_TASK_PREFIX L"%s", sid);
    context->sid = SysAllocString(sid);
    LocalFree(sid);
    if (count <= 0 || count >= (int)(sizeof(name) / sizeof(*name))) return ERROR_INSUFFICIENT_BUFFER;
    context->name = SysAllocString(name);
    if (!context->sid || !context->name) return ERROR_NOT_ENOUGH_MEMORY;
    result = CoCreateInstance(&CLSID_TaskScheduler, NULL, CLSCTX_INPROC_SERVER,
                              &IID_ITaskService, (void **)&context->service);
    if (FAILED(result)) return TaskError(result);
    VariantInit(&empty);
    result = ITaskService_Connect(context->service, empty, empty, empty, empty);
    if (FAILED(result)) return TaskError(result);
    root = SysAllocString(L"\\");
    if (!root) return ERROR_NOT_ENOUGH_MEMORY;
    result = ITaskService_GetFolder(context->service, root, &context->folder);
    SysFreeString(root);
    return TaskError(result);
}

/* Task Scheduler can return an account name even when registered with a SID. */
static BOOL StartupTaskUserMatches(const wchar_t *user, const wchar_t *expected)
{
    BYTE sid[SECURITY_MAX_SID_SIZE];
    wchar_t domain[256], *actual = NULL;
    DWORD sidBytes = sizeof(sid), domainChars = sizeof(domain) / sizeof(*domain);
    SID_NAME_USE kind;
    if (!user) return FALSE;
    if (!_wcsicmp(user, expected)) return TRUE;
    if (!LookupAccountNameW(NULL, user, sid, &sidBytes, domain, &domainChars, &kind) ||
        !ConvertSidToStringSidW(sid, &actual)) return FALSE;
    BOOL matches = !_wcsicmp(actual, expected);
    LocalFree(actual);
    return matches;
}

/* A task for another copy is not this copy's enabled setting. Inspect its
   actual action rather than a cached registry flag; Task Scheduler can disable it. */
static HRESULT StartupTaskMatches(IRegisteredTask *task, const StartupTaskContext *context,
                                  BOOL *matches)
{
    ITaskDefinition *definition = NULL;
    IActionCollection *actions = NULL;
    IAction *action = NULL;
    IExecAction *execute = NULL;
    IPrincipal *principal = NULL;
    BSTR path = NULL, arguments = NULL, user = NULL;
    LONG count = 0;
    TASK_RUNLEVEL_TYPE level = TASK_RUNLEVEL_LUA;
    TASK_LOGON_TYPE logon = TASK_LOGON_NONE;
    HRESULT result = IRegisteredTask_get_Definition(task, &definition);
    *matches = FALSE;
    if (SUCCEEDED(result)) result = ITaskDefinition_get_Actions(definition, &actions);
    if (SUCCEEDED(result)) result = IActionCollection_get_Count(actions, &count);
    if (SUCCEEDED(result) && count != 1) result = S_FALSE;
    if (result == S_OK) result = IActionCollection_get_Item(actions, 1, &action);
    if (result == S_OK) result = IAction_QueryInterface(action, &IID_IExecAction, (void **)&execute);
    if (result == S_OK) result = IExecAction_get_Path(execute, &path);
    if (result == S_OK) result = IExecAction_get_Arguments(execute, &arguments);
    if (result == S_OK) result = ITaskDefinition_get_Principal(definition, &principal);
    if (result == S_OK) result = IPrincipal_get_RunLevel(principal, &level);
    if (result == S_OK) result = IPrincipal_get_LogonType(principal, &logon);
    if (result == S_OK) result = IPrincipal_get_UserId(principal, &user);
    if (result == S_OK)
        *matches = path && !_wcsicmp(path, context->path) && (!arguments || !*arguments) &&
            StartupTaskUserMatches(user, context->sid) && level == TASK_RUNLEVEL_HIGHEST &&
            logon == TASK_LOGON_INTERACTIVE_TOKEN;
    SysFreeString(path);
    SysFreeString(arguments);
    SysFreeString(user);
    if (principal) IPrincipal_Release(principal);
    if (execute) IExecAction_Release(execute);
    if (action) IAction_Release(action);
    if (actions) IActionCollection_Release(actions);
    if (definition) ITaskDefinition_Release(definition);
    return result;
}

LONG ReadStartupTask(BOOL *present, BOOL *enabled)
{
    StartupTaskContext context = {0};
    IRegisteredTask *task = NULL;
    *present = *enabled = FALSE;
    LONG error = OpenStartupTask(&context);
    if (error == ERROR_SUCCESS) {
        HRESULT result = ITaskFolder_GetTask(context.folder, context.name, &task);
        if (result != HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND)) {
            if (SUCCEEDED(result)) {
                BOOL matches = FALSE;
                VARIANT_BOOL active = VARIANT_FALSE;
                *present = TRUE;
                result = StartupTaskMatches(task, &context, &matches);
                if (SUCCEEDED(result)) result = IRegisteredTask_get_Enabled(task, &active);
                if (SUCCEEDED(result)) *enabled = matches && active != VARIANT_FALSE;
            }
            error = TaskError(result);
        }
    }
    if (task) IRegisteredTask_Release(task);
    CloseStartupTask(&context);
    return error;
}

static BOOL EscapeTaskXml(const wchar_t *source, wchar_t *output, size_t capacity)
{
    size_t used = 0;
    while (*source) {
        const wchar_t *escaped = *source == L'&' ? L"&amp;" :
            *source == L'<' ? L"&lt;" : *source == L'>' ? L"&gt;" : NULL;
        wchar_t single[2] = {*source++, 0};
        if (!escaped) escaped = single;
        size_t count = wcslen(escaped);
        if (count >= capacity - used) return FALSE;
        memcpy(output + used, escaped, count * sizeof(*output));
        used += count;
    }
    output[used] = 0;
    return TRUE;
}

LONG WriteStartupTask(BOOL enable)
{
    StartupTaskContext context = {0};
    IRegisteredTask *task = NULL;
    LONG error = OpenStartupTask(&context);
    HRESULT result = S_OK;
    if (error != ERROR_SUCCESS) goto done;
    if (!enable) {
        result = ITaskFolder_DeleteTask(context.folder, context.name, 0);
        if (result == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND)) result = S_OK;
    } else {
        wchar_t path[MAX_PATH * 5], xml[4096];
        VARIANT empty, user;
        VariantInit(&empty);
        VariantInit(&user);
        user.vt = VT_BSTR;
        user.bstrVal = context.sid;
        if (!EscapeTaskXml(context.path, path, sizeof(path) / sizeof(*path))) {
            error = ERROR_INSUFFICIENT_BUFFER;
            goto done;
        }
        /* InteractiveToken stores no password and never runs as SYSTEM. No
           time/battery limit may terminate a tray application after sign-in. */
        int count = swprintf(xml, sizeof(xml) / sizeof(*xml),
            L"<Task version=\"1.2\" xmlns=\"http://schemas.microsoft.com/windows/2004/02/mit/task\">"
            L"<Triggers><LogonTrigger><Enabled>true</Enabled><UserId>%s</UserId></LogonTrigger></Triggers>"
            L"<Principals><Principal id=\"User\"><UserId>%s</UserId>"
            L"<LogonType>InteractiveToken</LogonType><RunLevel>HighestAvailable</RunLevel></Principal></Principals>"
            L"<Settings><MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>"
            L"<DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>"
            L"<StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>"
            L"<StartWhenAvailable>true</StartWhenAvailable><Enabled>true</Enabled>"
            L"<ExecutionTimeLimit>PT0S</ExecutionTimeLimit></Settings>"
            L"<Actions Context=\"User\"><Exec><Command>%s</Command></Exec></Actions></Task>",
            context.sid, context.sid, path);
        if (count <= 0 || count >= (int)(sizeof(xml) / sizeof(*xml))) {
            error = ERROR_INSUFFICIENT_BUFFER;
            goto done;
        }
        BSTR definition = SysAllocString(xml);
        if (!definition) { error = ERROR_NOT_ENOUGH_MEMORY; goto done; }
        result = ITaskFolder_RegisterTask(context.folder, context.name, definition,
            TASK_CREATE_OR_UPDATE, user, empty, TASK_LOGON_INTERACTIVE_TOKEN, empty, &task);
        SysFreeString(definition);
    }
    error = TaskError(result);
done:
    if (task) IRegisteredTask_Release(task);
    CloseStartupTask(&context);
    return error;
}
