# Override Outlook Document Theme Add-In

This simple VSTO Outlook Add-In allows you to force Outlook to use a predefined theme for new emails.
The Add-In was born out of an issue at a customer who didn't want to use the "new" Aptos Theme and there was no easy way to revert this change.

The Add-In hooks into the ItemOpen event for unsent WordEditor emails. It then changes the active theme. By default it will change to "Office 2013 - 2022 Theme.thmx".
You can specify a different theme by creating a registry-key under "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\MailSettings" called "OverrideTheme". The Add-In will then look for the theme in "{OfficeInstallationPath}\root\Document Themes 16".