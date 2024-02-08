# Override Outlook Document Theme Add-In

This simple VSTO Outlook Add-In allows you to force Outlook 365 to use a predefined theme for new emails.\
The Add-In was born out of an issue at a customer, who didn't want to use the new Aptos Theme and we couldn't find an easier/nicer way to revert this change.

The Add-In hooks into the [ItemOpen-event](https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.itemevents_10_event.open) for unsent [WordEditor](https://learn.microsoft.com/en-us/office/vba/api/outlook.inspector.wordeditor) emails and changes its active theme.\
By default it will change to `Office 2013 - 2022 Theme.thmx`.

You can specify a different theme by creating a registry-key under `HKCU\Software\Microsoft\Office\16.0\Common\MailSettings` called `OverrideTheme`. The Add-In will then look for the theme in `{OfficeInstallationPath}\root\Document Themes 16`.
