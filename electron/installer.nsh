!macro customInstall
  ; Create .env file from template if it doesn't exist
  IfFileExists "$INSTDIR\resources\backend\.env" env_exists env_missing
  env_missing:
    CopyFiles "$INSTDIR\resources\backend\.env.template" "$INSTDIR\resources\backend\.env"
  env_exists:

  ; Create user data directory
  CreateDirectory "$APPDATA\MailScraper"
  CreateDirectory "$APPDATA\MailScraper\attachments"
  CreateDirectory "$APPDATA\MailScraper\exports"

  ; Write a first-run marker
  FileOpen $0 "$APPDATA\MailScraper\.first_run" w
  FileClose $0

!macroend

!macro customUninstall
  ; Ask user if they want to delete their data
  MessageBox MB_YESNO "Do you want to delete all your scraped email data and candidate profiles? $\n$\nChoose No to keep your data for future reinstalls." IDYES delete_data IDNO keep_data
  delete_data:
    RMDir /r "$APPDATA\MailScraper"
  keep_data:
!macroend
