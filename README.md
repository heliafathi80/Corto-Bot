# Corto-Bot
Below is a consolidated requirements document based on everything you’ve provided so far.

# Documents Sync Bot — High-Level Requirements

## 1. Purpose

The Documents Sync Bot is a Windows desktop application that provides **two-way synchronization between local files/folders and the server/cloud**, similar in concept to Dropbox or OneDrive.

The system must allow users to work with matter-related documents locally while ensuring changes are synchronized with the central application.

The solution includes three locally running components:

1. **Main Web Application / PWA**
2. **Microsoft Word / Outlook COM Add-in**
3. **Documents Sync Bot**

The Documents Sync Bot acts as the local synchronization and orchestration component.

---

# 2. High-Level Architecture

```text
                    Backend / Cloud
                          ▲
                          │ HTTPS
                          │
                 ┌────────┴────────┐
                 │ Documents Sync  │
                 │      Bot        │
                 │                 │
                 │ Sync Engine     │
                 │ SQLite          │
                 │ File Watcher    │
                 │ Local API       │
                 └───────┬─────────┘
                         │
            ┌────────────┴────────────┐
            │                         │
      Local HTTP/API             Named Pipes
            │                         │
            ▼                         ▼
       Main Web App             COM Add-in
          / PWA              Word / Outlook
```

The backend/cloud is the authoritative source for shared document state.

---

# 3. Documents Bot

## 3.1 Application Type

The Documents Bot should be implemented as a **Windows desktop application**, preferably:

```text
.NET Windows Forms application
```

The application may run primarily as a **system tray application**.

A Windows Service is not required for the initial implementation.

The bot should run in the context of the currently logged-in Windows user.

---

## 3.2 Primary Responsibilities

The bot must:

- communicate with the main web application
- communicate with Word/Outlook COM add-ins
- maintain local matter-folder mappings
- monitor mapped folders for filesystem changes
- synchronize files from local storage to the server
- synchronize server changes back to the local machine
- identify potential matter folders
- manage authentication context
- support multiple application users/accounts
- detect sync conflicts
- persist sync metadata locally
- retry failed synchronization operations
- notify the application about relevant sync events

---

# 4. Local Inter-Process Communication

Two different local communication mechanisms are required.

## 4.1 COM Add-in → Documents Bot

Communication between the Word/Outlook COM add-in and the bot should use:

```text
Windows Named Pipes
```

The Documents Bot should act as the named-pipe server.

The COM add-in should act as the client.

Example:

```text
Word
  ↓
COM Add-in
  ↓ Named Pipe
Documents Bot
```

Named-pipe communication should support bidirectional messaging where required.

---

## 4.2 Main Web App / PWA → Documents Bot

A browser cannot directly communicate with Windows Named Pipes.

The Documents Bot should therefore expose a local HTTP API bound only to localhost.

Example:

```text
http://127.0.0.1:<port>
```

Communication:

```text
PWA
 ↓
localhost HTTP
 ↓
Documents Bot
```

Example endpoint:

```http
POST /api/notifications/matter-folder-import
```

The local API must not be exposed externally.

---

## 4.3 Bot → PWA Notifications

If the bot needs to notify the PWA asynchronously, the solution should support a real-time local communication mechanism such as:

```text
WebSocket
```

or:

```text
SignalR
```

This should only be required where simple request/response HTTP is insufficient.

---

# 5. Common Message Model

All messages received by the Documents Bot should ultimately be processed through a shared message-dispatching layer.

Example:

```text
Named Pipe ─────┐
                │
                ▼
           MessageDispatcher
                ▲
                │
Local HTTP ─────┘
```

Communication transport must not contain business logic.

---

# 6. COM Add-in User Journey

When a user opens a document in Microsoft Word:

```text
User opens document
      ↓
COM add-in is activated
      ↓
COM add-in determines local document path
      ↓
COM add-in notifies Documents Bot
```

Notification type:

```text
documentLocation
```

Example payload:

```json
{
  "type": "documentLocation",
  "documentPath":
    "C:\\Users\\username\\Documents\\Matters\\JohnDoe-Purchase-Matter\\contracts\\contract.docx"
}
```

The bot should then process the document location.

---

# 7. `documentLocation` Processing

When the bot receives a `documentLocation` notification:

## Step 1

Determine the document's parent directory.

Example:

```text
C:\Documents\Matters\Smith\contracts\contract.docx
```

Parent:

```text
C:\Documents\Matters\Smith\contracts
```

---

## Step 2

Generate a snapshot of the directory.

The snapshot should include relevant information such as:

- child directory names
- document names
- file metadata where required

Existing cloud-storage-provider classification logic should be reused where practical.

---

## Step 3

Send the snapshot to a folder-classification endpoint.

The server determines whether the folder appears to represent a matter folder.

The existing classifier currently accepts cloud-storage folder IDs.

A new interface/endpoint may be required that can classify based on a local folder snapshot.

---

## Step 4 — Matter Folder Detected

If classified as a matter folder:

- create a matter-folder record
- set its status to something equivalent to:

```text
suggested
```

- notify the front end using the existing CSP-style mechanism where possible

---

## Step 5 — Not a Matter Folder

If the directory is not classified as a matter folder:

```text
move one directory level upward
```

Generate another snapshot and repeat classification.

Example:

```text
...\Smith\Contracts
       ↓
...\Smith
       ↓
...\Matters
```

The search should terminate at an appropriate configured/root boundary.

---

# 8. Main Web Application User Journey

A user may explicitly import a local matter directory.

Example:

```text
Matter
 ↓
Import from local drive
 ↓
Select folder
 ↓
Main application informs Documents Bot
```

Notification type:

```text
matterFolderImport
```

Example:

```json
{
  "type": "matterFolderImport",
  "matterId": "123",
  "matterFolderPath":
    "C:\\Users\\username\\Documents\\Matters\\JohnDoe-Purchase-Matter"
}
```

---

# 9. `matterFolderImport` Processing

When receiving this notification, the bot should:

## 9.1 Store Matter Mapping

Persist the mapping between:

```text
MatterId
↕
LocalFolderPath
```

in local SQLite storage.

---

## 9.2 Initial Upload

Synchronize the contents of the selected local directory to the cloud/server.

---

## 9.3 Inspect Neighboring Folders

Move one directory level up.

Example:

```text
C:\Documents\Matters\Matter123

→

C:\Documents\Matters
```

Enumerate sibling directories.

---

## 9.4 Classify Potential Matter Folders

For each sibling directory:

1. generate folder snapshot
2. call classification service
3. determine whether the directory is likely a matter folder

If yes:

```text
create suggested matter-folder entry
```

If not:

```text
continue to next directory
```

---

# 10. Two-Way Synchronization

The primary requirement is **continuous two-way synchronization**.

The bot must support:

```text
Local → Server
Server → Local
```

It is not simply an upload application.

---

# 11. Local File Monitoring

Mapped directories should be monitored for changes.

The implementation may use:

```text
FileSystemWatcher
```

but filesystem events must not be considered fully authoritative.

A reconciliation mechanism should also periodically verify filesystem state.

Relevant local changes include:

- file creation
- file update
- file deletion
- file rename
- directory creation
- directory rename
- directory deletion

---

# 12. Server-Side Change Detection

The bot must be informed when the server representation of a matter changes.

Potential mechanisms include:

- SignalR
- WebSockets
- server notifications
- polling as a fallback

Server-originated changes should trigger local synchronization.

---

# 13. Multi-User Support

The bot must support **multiple authenticated application users**.

Multiple users may use the same Windows machine.

Example:

```text
Windows account
    │
    ├── Application User A
    └── Application User B
```

The bot must therefore not assume:

```text
Windows user == application user
```

---

# 14. Account Context

The bot should maintain separate account contexts.

Example:

```text
Documents Bot
│
├── AccountContext A
│
├── AccountContext B
│
└── AccountContext C
```

Each account context may contain:

- user identifier
- tenant/organisation identifier
- authentication state
- access credentials/reference
- matter mappings
- sync jobs
- server connection state

---

# 15. Authentication

The bot must not trust an account ID supplied by an IPC message alone.

A client should establish an authenticated local session.

Example:

```text
PWA
 ↓
authenticated with backend
 ↓
register/connect with bot
 ↓
bot validates identity
 ↓
BotSession created
```

Subsequent messages should be associated with the authenticated session.

---

# 16. Bot Sessions

A bot session should conceptually contain:

```text
SessionId
AccountId
TenantId
Source
CreatedAt
ExpiresAt
```

Possible sources:

```text
PWA
Word COM Add-in
Outlook COM Add-in
```

---

# 17. COM Add-in Authentication

The COM add-in must also establish an identity/context with the bot.

When a named-pipe connection is established, the bot should associate that connection with a validated application account.

Once authenticated:

```text
NamedPipe Connection 12
        ↓
User A
Tenant X
```

Subsequent messages received from that connection can inherit the corresponding identity.

---

# 18. Multiple Windows Users

Where multiple Windows login sessions exist:

```text
Windows User Alice
Windows User Bob
```

each logged-in Windows session should run its own instance of the Documents Bot.

Example:

```text
Alice Windows session
    └── Documents Bot

Bob Windows session
    └── Documents Bot
```

A single global Windows Service should not be responsible for all users' interactive file synchronization.

---

# 19. Local Storage

SQLite should be used for local metadata/state.

It should not be treated as the authoritative document store.

---

# 20. Matter Folder Mapping

A mapping should contain at least:

```text
AccountId
TenantId
MatterId
LocalFolderPath
SyncStatus
```

Mappings must be account-aware.

A `MatterId` alone must not uniquely identify a local mapping.

---

# 21. File Sync State

The bot should persist per-file synchronization metadata.

Example structure:

```text
SyncedFile
--------------------------
AccountId
TenantId
MatterId
RemoteFileId
RemoteFolderId
LocalPath
RemoteVersion
LocalHash
LastSyncedHash
LastSyncedAtUtc
SyncStatus
```

---

# 22. Stable File Identity

File paths must not be treated as file identity.

Each cloud/server file should have a stable identifier:

```text
RemoteFileId
```

Similarly, server folders should preferably have:

```text
RemoteFolderId
```

This allows a file to be renamed without being treated as:

```text
delete old file
+
create new file
```

---

# 23. Concurrent Users Editing the Same Matter

Multiple users may synchronize the same matter simultaneously.

Example:

```text
User A → Matter 123
User B → Matter 123
```

The server must be the source of truth for resolving concurrent changes.

Bots should not synchronize directly with each other.

Correct architecture:

```text
Bot A ──┐
        ├── Server
Bot B ──┘
```

Not:

```text
Bot A ↔ Bot B
```

---

# 24. File-Level Concurrency

Concurrency should be detected at the individual file level.

Example:

```text
Matter 123
├── Contract.docx
├── Notes.docx
└── Invoice.pdf
```

If:

```text
User A changes Contract.docx
User B changes Notes.docx
```

both changes should sync normally.

A conflict exists when both modify the same logical file.

---

# 25. Server-Side Versioning

Every server-side file should have a version or equivalent concurrency token.

Example:

```text
FileId
Version
Hash
LastModifiedUtc
UpdatedBy
```

Example:

```text
Contract.docx
Version 15
```

The client should upload with:

```text
ExpectedVersion = 15
```

The server must atomically check whether version 15 is still current.

---

# 26. Optimistic Concurrency

The synchronization system should use optimistic concurrency.

Example:

```text
User A downloads version 15
User B downloads version 15

A uploads
→ version becomes 16

B tries uploading based on version 15
→ server detects conflict
```

The server should respond with an appropriate concurrency error such as:

```http
409 Conflict
```

or:

```http
412 Precondition Failed
```

when ETags are used.

---

# 27. ETag Support

The API may use HTTP ETags.

Example upload:

```http
PUT /matters/{matterId}/files/{fileId}

If-Match: "version-15"
```

Server response:

```http
ETag: "version-16"
```

This allows atomic concurrency checking.

---

# 28. Conflict Handling

The bot must never silently overwrite another user's changes.

For an initial implementation:

```text
Local changed + Remote unchanged
→ upload local

Local unchanged + Remote changed
→ download remote

Local changed + Remote changed
→ conflict

Neither changed
→ no action
```

---

# 29. Conflict Copies

When the same file changes independently on multiple clients, both versions should be preserved.

Example:

```text
Contract.docx

Contract (Amir's conflicted copy 2026-09-29).docx
```

The user should be notified about the conflict.

Automatic merging of Microsoft Word documents is not required for the initial implementation.

---

# 30. Delete/Edit Conflicts

Example:

```text
User A deletes Contract.docx
User B modifies Contract.docx
```

The system should avoid silently deleting User B's work.

Recommended initial behaviour:

```text
preserve the modified document
+
mark conflict
```

---

# 31. Rename Handling

A rename should maintain the same server file identity.

Example:

Before:

```text
FileId 456
Contract.docx
```

After:

```text
FileId 456
Final Contract.docx
```

The server should understand this as a rename rather than:

```text
delete + create
```

where possible.

---

# 32. Synchronization Queue

File operations should be processed using a durable sync queue.

Each operation should contain sufficient identity information.

Example:

```text
SyncJob
----------------------
AccountId
TenantId
MatterId
RemoteFileId
LocalPath
Operation
ExpectedVersion
RetryCount
CreatedAt
```

Operations may include:

```text
Upload
Download
Delete
Rename
CreateFolder
DeleteFolder
```

---

# 33. Retry Behaviour

Temporary failures should not immediately result in permanent sync failure.

The bot should support:

- retry with backoff
- retry count
- transient-error detection
- network reconnect
- offline operation

Failed jobs should remain available for retry.

---

# 34. Offline Operation

The user may modify files while disconnected from the server.

The bot should:

```text
detect change
↓
record pending operation locally
↓
wait for connectivity
↓
synchronize later
```

The concurrency/version check must still occur before upload.

---

# 35. Idempotency

Synchronization operations must be designed to tolerate retries.

Repeating an operation should not unintentionally create duplicate files or records.

Where required, operations should contain an idempotency key.

---

# 36. File Hashing

The bot should use file hashes where useful to determine whether file content actually changed.

Example:

```text
LocalHash
LastSyncedHash
```

This prevents unnecessary uploads caused by metadata-only filesystem events.

---

# 37. Change Detection State

A useful local state model is:

```text
LocalHash
LastSyncedHash
RemoteVersion
LastSyncedRemoteVersion
```

This allows determination of:

```text
LocalChanged
RemoteChanged
```

independently.

---

# 38. Avoiding Synchronization Loops

The bot must avoid the following scenario:

```text
server change
→ bot downloads
→ FileSystemWatcher detects download
→ bot uploads same file
→ server sends change again
→ ...
```

The bot must be able to recognise filesystem changes generated by its own synchronization operations.

---

# 39. Security — Local API

The local HTTP server must:

- bind only to loopback
- not expose itself to the LAN
- authenticate/authorise callers
- validate origins
- validate all input
- prevent arbitrary filesystem access
- use unpredictable/session-specific credentials where appropriate

---

# 40. Security — Named Pipes

Named pipes should:

- use appropriate Windows ACLs
- restrict access to the current Windows user/session where practical
- validate incoming messages
- reject malformed or unauthenticated clients

---

# 41. Credential Storage

Raw authentication tokens must not be stored unencrypted in SQLite.

Windows-protected credential storage should be used.

Possible approaches include:

```text
Windows Credential Manager
```

or:

```text
DPAPI-protected storage
```

SQLite may store a reference to the protected credential.

---

# 42. Suggested Bot Project Structure

A possible Visual Studio solution structure:

```text
Corto.DocumentsBot
│
├── Program.cs
├── TrayApplicationContext.cs
│
├── Communication
│   ├── NamedPipeServer.cs
│   ├── LocalApiServer.cs
│   ├── WebSocketServer.cs
│   └── MessageDispatcher.cs
│
├── Authentication
│   ├── AccountContext.cs
│   ├── BotSession.cs
│   ├── SessionManager.cs
│   └── CredentialStore.cs
│
├── Sync
│   ├── SyncEngine.cs
│   ├── LocalFileWatcher.cs
│   ├── ReconciliationService.cs
│   ├── SyncQueue.cs
│   ├── UploadProcessor.cs
│   ├── DownloadProcessor.cs
│   └── ConflictResolver.cs
│
├── Classification
│   ├── FolderSnapshotService.cs
│   └── FolderClassifier.cs
│
├── Storage
│   ├── BotDbContext.cs
│   ├── MatterFolderMapping.cs
│   ├── SyncedFile.cs
│   └── SyncJob.cs
│
├── Cloud
│   ├── DocumentsApiClient.cs
│   └── ServerNotificationClient.cs
│
└── Messages
    ├── BotMessage.cs
    ├── DocumentLocationMessage.cs
    └── MatterFolderImportMessage.cs
```

---

# 43. Initial POC Scope

For the first POC, I would define the minimum implementation as:

1. Windows tray application starts successfully.
2. SQLite database is created.
3. COM add-in can connect to the bot via Named Pipes.
4. COM add-in can send `documentLocation`.
5. PWA can connect to the bot through localhost HTTP.
6. PWA can send `matterFolderImport`.
7. Bot can associate messages with an authenticated user.
8. Bot can store:

```text
Account + MatterId + LocalFolder
```

9. Bot can monitor a mapped directory.
10. Bot detects new/modified files.
11. Bot can upload them to a test/backend API.
12. Bot can receive/download a server-side file change.
13. Basic server file versioning exists.
14. Concurrent updates return a conflict instead of overwriting.
15. Conflicting local versions are preserved.
16. Folder snapshots can be generated.
17. Folder classification endpoint can be called.
18. Suggested matter folders can be created.

That POC would validate the major architectural risks before implementing full Dropbox-style synchronization.

---

# 44. Key Architectural Principles

The system should follow these principles:

```text
Server is authoritative for shared state.

Bot is authoritative for observing the local filesystem.

File identity != file path.

Every sync action belongs to an authenticated account.

Concurrency is resolved by server-side version checking.

Never silently overwrite conflicting changes.

Transport and business logic remain separate.

Sync state must survive application restarts.

Local changes must survive temporary loss of connectivity.

Multiple users may operate on the same matter concurrently.
```
