# Voximplant API client for MS Excel VBA
## Requirements
- Developer toolbar enabled in Excel
- Microsoft XML, v6.0 and Microsoft Scripting Runtime references enabled in Tools->References
- JSON parsing module installed from https://github.com/VBA-tools/VBA-JSON
- Voximplant account
## Installation
1. Download VBA file
1. Add a new Class Module and rename it into VoximplantAPI
1. Paste code from VBA file into that module
## Basic example
```
    Dim res As Dictionary
    Dim api As New VoximplantAPI
    ' Pass acocunt id and API key here
    api.SetCredentials 1, "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa"
    Set res = api.GetAccountInfo
```