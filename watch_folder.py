from enum import Enum, auto
import win32file
import win32con


class FileChangeActions(Enum):
    CREATED = auto()
    DELETED = auto()
    UPDATED = auto()
    RENAMED = auto()

def watch_folder(watch_path: str, ignore_extraneous_actions: bool=True) -> None:
    watch_dir = win32file.CreateFile(
        watch_path,
        win32file.GENERIC_READ,
        (
            win32con.FILE_SHARE_READ |
            win32con.FILE_SHARE_WRITE |
            win32con.FILE_SHARE_DELETE
        ),
        None,
        win32con.OPEN_EXISTING,
        win32con.FILE_FLAG_BACKUP_SEMANTICS,
        None
    )

    while True:
        file_change = win32file.ReadDirectoryChangesW(
            watch_dir,
            1024,
            True,
            (
                win32con.FILE_NOTIFY_CHANGE_FILE_NAME |
                win32con.FILE_NOTIFY_CHANGE_DIR_NAME |
                win32con.FILE_NOTIFY_CHANGE_ATTRIBUTES |
                win32con.FILE_NOTIFY_CHANGE_SIZE |
                win32con.FILE_NOTIFY_CHANGE_LAST_WRITE |
                win32con.FILE_NOTIFY_CHANGE_SECURITY
            ),
            None,
            None
        )

        file_change = dict(file_change)

        if file_change.keys() in ({1}, {2}, {3}):
            for change_action in file_change:
                file_change = {
                    'action': FileChangeActions(change_action),
                    'file': file_change[change_action]
                }
        elif file_change.keys() == {4, 5}:
            file_change = {
                'action': FileChangeActions.RENAMED,
                'from': file_change[4],
                'to': file_change[5]
            }
        else:
            if not ignore_extraneous_actions:
                actions = ', '.join(str(r) for r in file_change)

                raise ValueError(
                    f'The file watcher received extraneous actions: {actions}'
                )

        print(file_change)
