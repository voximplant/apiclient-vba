# Voximplant API client for MS Excel VBA
## Requirements
- Developer toolbar enabled in Excel
- Microsoft XML, v6.0 and Microsoft Scripting Runtime references enabled in Tools->References
- JSON parsing module installed from https://github.com/VBA-tools/VBA-JSON
- Voximplant account
## Basic example
```
    Dim res As Dictionary
    Dim api As New VoximplantAPI
    ' Pass acocunt id and API key here
    api.SetCredentials 1, "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa"
    Set res = api.GetAccountInfo
```