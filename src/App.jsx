M

mmbuwi-crypto's Project

main

Connect

Help

Advisor Center

SQL Editor

AI Assistant

Edge Functionssvg]:size-3.5 text-foreground-muted" style="box-sizing: border-box; border-width: 0px; border-style: solid; border-color: var(--border-default,currentColor); border-image: initial; margin: 0px; padding: 0px; color: var(--foreground-muted); font-size: var(--text-xs); line-height: var(--tw-leading,var(--text-xs--line-height));">manage-teacher

manage-teacher

https://xapbkapxpdvdvelcbpyo.supabase.co/functions/v1/manage-teacher

a few seconds ago

Docs

Download

Test

OverviewInvocationsLogsCodeSettings

Function configuration

Name

Your slug and endpoint URL will remain the same

Verify JWT with legacy secret

Requires a JWT signed only by the legacy secret in the Authorization header. The anon key satisfies this.

Recommended: OFF with JWT and custom auth logic in your function code.

Save changes

Invoke function

cURLJavaScriptSwiftFlutterPythonShow anon key

code]:wrap-break-word" contenteditable="true" style="box-sizing: border-box; border-width: 1px; border-style: none; border-color: var(--border-default,currentColor); border-image: initial; margin-right: 0px; margin-bottom: 0px; margin-left: 0px; padding: calc(var(--spacing) * 0); font-family: var(--default-mono-font-family,ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace); font-feature-settings: var(--default-mono-font-feature-settings,normal); font-variation-settings: var(--default-mono-font-variation-settings,normal); font-size: 13px; margin-block: calc(var(--spacing) * 0) !important; margin-top: calc(var(--spacing) * 0) !important; width: 286.933px; border-radius: var(--radius-md); --tw-border-style: none; background-color: var(--background-surface-100) !important; line-height: 1.4; --tw-outline-style: none; outline-style: none; display: block; overflow-x: auto; color: rgb(68, 68, 68);">curl -L -X POST 'https://xapbkapxpdvdvelcbpyo.supabase.co/functions/v1/manage-teacher' \ -H 'Authorization: Bearer SUPABASE_PUBLISHABLE_KEY' \ -H 'apikey: SUPABASE_PUBLISHABLE_KEY' \ -H 'Content-Type: application/json' \ --data '{"name":"Functions"}'

Develop locally

> 1. Download the function

$supabase functions download manage-teacher 

> Deploy a new version

$supabase functions deploy manage-teacher 

> Delete the function

