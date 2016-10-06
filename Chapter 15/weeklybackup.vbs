'weeklybackup.vbs
Const MD_BACKUP_NEXT_VERSION = &HFFFFFFFF
Const MD_BACKUP_FORCE_BACKUP = 4
Const MD_BACKUP_SAVE_FIRST = 2

Set objComputer = GetObject("IIS://odin")

'create a new backup. Assign a new version number
objComputer.Backup "Weekly backup", MD_BACKUP_NEXT_VERSION _
       , MD_BACKUP_FORCE_BACKUP Or MD_BACKUP_SAVE_FIRST
