Private api_key As String
Private account_id As Integer

Private Function URL_Encode(ByRef txt As String) As String
    Dim buffer As String, i As Long, c As Long, n As Long
    buffer = String$(Len(txt) * 12, "%")

    For i = 1 To Len(txt)
        c = AscW(Mid$(txt, i, 1)) And 65535

        Select Case c
            Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95  ' Unescaped 0-9A-Za-z-._ '
                n = n + 1
                Mid$(buffer, n) = ChrW(c)
            Case Is <= 127            ' Escaped UTF-8 1 bytes U+0000 to U+007F '
                n = n + 3
                Mid$(buffer, n - 1) = Right$(Hex$(256 + c), 2)
            Case Is <= 2047           ' Escaped UTF-8 2 bytes U+0080 to U+07FF '
                n = n + 6
                Mid$(buffer, n - 4) = Hex$(192 + (c \ 64))
                Mid$(buffer, n - 1) = Hex$(128 + (c Mod 64))
            Case 55296 To 57343       ' Escaped UTF-8 4 bytes U+010000 to U+10FFFF '
                i = i + 1
                c = 65536 + (c Mod 1024) * 1024 + (AscW(Mid$(txt, i, 1)) And 1023)
                n = n + 12
                Mid$(buffer, n - 10) = Hex$(240 + (c \ 262144))
                Mid$(buffer, n - 7) = Hex$(128 + ((c \ 4096) Mod 64))
                Mid$(buffer, n - 4) = Hex$(128 + ((c \ 64) Mod 64))
                Mid$(buffer, n - 1) = Hex$(128 + (c Mod 64))
            Case Else                 ' Escaped UTF-8 3 bytes U+0800 to U+FFFF '
                n = n + 9
                Mid$(buffer, n - 7) = Hex$(224 + (c \ 4096))
                Mid$(buffer, n - 4) = Hex$(128 + ((c \ 64) Mod 64))
                Mid$(buffer, n - 1) = Hex$(128 + (c Mod 64))
        End Select
    Next
    URL_Encode = Left$(buffer, n)
End Function



Public Sub SetCredentials(accountId As Integer, ByRef apiKey As String)
    api_key = apiKey
    account_id = accountId
End Sub

Public Function makeRequest(name As String, params As Dictionary) As Object

    Dim objHTTP As New MSXML2.XMLHTTP60
    Dim jsonData As String
    Dim parsedJson As Object
    Dim postString As String


    postString = ""

    Dim iterKey As Variant

    For Each iterKey In params.Keys
        postString = postString & "&" & iterKey & "=" & URL_Encode(params(iterKey))
    Next


    Debug.Print(postString)
    Url = "https://api.voximplant.com/platform_api/" + name
    objHTTP.Open "POST", Url, False
    objHTTP.send "account_id=" & account_id & "&api_key=" & api_key & postString
    jsonData = objHTTP.responseText
    Debug.Print(jsonData)
    Set parsedJson = JsonConverter.ParseJson(jsonData)
    Set makeRequest = parsedJson



End Function

Public Function vba_datetime_to_api(dt) As String
    vba_datetime_to_api = Format(dt, "yyyy-mm-dd hh:mm:ss")
End Function

Public Function vba_date_to_api(dt) As String
    vba_date_to_api = Format(dt, "yyyy-mm-dd")
End Function



Public Function serialize_list(l) As String
    If IsArray(l) Then
        Dim x As Variant
        Dim r As String
        r = ""
        For Each x In l
            If r = "" Then
                r = x
            Else
                r = r & ";" & x
            End If
        Next
        serialize_list = r
    Else
        serialize_list = l
    End If

End Function

Public Function GetAccountInfo(Optional return_live_balance = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(return_live_balance) Then
        params.Add "return_live_balance", return_live_balance

    End If

    Set res = makeRequest("GetAccountInfo", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetAccountInfo = res
End Function

Public Function SetAccountInfo(Optional new_account_email = Null, Optional new_account_password = Null, Optional language_code = Null, Optional location = Null, Optional account_first_name = Null, Optional account_last_name = Null, Optional min_balance_to_notify = Null, Optional account_notifications = Null, Optional tariff_changing_notifications = Null, Optional news_notifications = Null, Optional send_js_error = Null, Optional billing_address_name = Null, Optional billing_address_country_code = Null, Optional billing_address_address = Null, Optional billing_address_zip = Null, Optional billing_address_phone = Null, Optional account_custom_data = Null, Optional callback_url = Null, Optional callback_salt = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(new_account_email) Then
        params.Add "new_account_email", new_account_email

    End If

    If Not IsNull(new_account_password) Then
        params.Add "new_account_password", new_account_password

    End If

    If Not IsNull(language_code) Then
        params.Add "language_code", language_code

    End If

    If Not IsNull(location) Then
        params.Add "location", location

    End If

    If Not IsNull(account_first_name) Then
        params.Add "account_first_name", account_first_name

    End If

    If Not IsNull(account_last_name) Then
        params.Add "account_last_name", account_last_name

    End If

    If Not IsNull(min_balance_to_notify) Then
        params.Add "min_balance_to_notify", min_balance_to_notify

    End If

    If Not IsNull(account_notifications) Then
        params.Add "account_notifications", account_notifications

    End If

    If Not IsNull(tariff_changing_notifications) Then
        params.Add "tariff_changing_notifications", tariff_changing_notifications

    End If

    If Not IsNull(news_notifications) Then
        params.Add "news_notifications", news_notifications

    End If

    If Not IsNull(send_js_error) Then
        params.Add "send_js_error", send_js_error

    End If

    If Not IsNull(billing_address_name) Then
        params.Add "billing_address_name", billing_address_name

    End If

    If Not IsNull(billing_address_country_code) Then
        params.Add "billing_address_country_code", billing_address_country_code

    End If

    If Not IsNull(billing_address_address) Then
        params.Add "billing_address_address", billing_address_address

    End If

    If Not IsNull(billing_address_zip) Then
        params.Add "billing_address_zip", billing_address_zip

    End If

    If Not IsNull(billing_address_phone) Then
        params.Add "billing_address_phone", billing_address_phone

    End If

    If Not IsNull(account_custom_data) Then
        params.Add "account_custom_data", account_custom_data

    End If

    If Not IsNull(callback_url) Then
        params.Add "callback_url", callback_url

    End If

    If Not IsNull(callback_salt) Then
        params.Add "callback_salt", callback_salt

    End If

    Set res = makeRequest("SetAccountInfo", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set SetAccountInfo = res
End Function

Public Function SetChildAccountInfo(Optional child_account_id = Null, Optional child_account_name = Null, Optional child_account_email = Null, Optional new_child_account_email = Null, Optional new_child_account_password = Null, Optional account_notifications = Null, Optional tariff_changing_notifications = Null, Optional news_notifications = Null, Optional active = Null, Optional language_code = Null, Optional location = Null, Optional min_balance_to_notify = Null, Optional support_robokassa = Null, Optional support_bank_card = Null, Optional support_invoice = Null, Optional can_use_restricted = Null, Optional min_payment_amount = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(child_account_id) Then
        params.Add "child_account_id", serialize_list(child_account_id)

    End If

    If Not IsNull(child_account_name) Then
        params.Add "child_account_name", serialize_list(child_account_name)

    End If

    If Not IsNull(child_account_email) Then
        params.Add "child_account_email", serialize_list(child_account_email)

    End If

    If Not IsNull(new_child_account_email) Then
        params.Add "new_child_account_email", new_child_account_email

    End If

    If Not IsNull(new_child_account_password) Then
        params.Add "new_child_account_password", new_child_account_password

    End If

    If Not IsNull(account_notifications) Then
        params.Add "account_notifications", account_notifications

    End If

    If Not IsNull(tariff_changing_notifications) Then
        params.Add "tariff_changing_notifications", tariff_changing_notifications

    End If

    If Not IsNull(news_notifications) Then
        params.Add "news_notifications", news_notifications

    End If

    If Not IsNull(active) Then
        params.Add "active", active

    End If

    If Not IsNull(language_code) Then
        params.Add "language_code", language_code

    End If

    If Not IsNull(location) Then
        params.Add "location", location

    End If

    If Not IsNull(min_balance_to_notify) Then
        params.Add "min_balance_to_notify", min_balance_to_notify

    End If

    If Not IsNull(support_robokassa) Then
        params.Add "support_robokassa", support_robokassa

    End If

    If Not IsNull(support_bank_card) Then
        params.Add "support_bank_card", support_bank_card

    End If

    If Not IsNull(support_invoice) Then
        params.Add "support_invoice", support_invoice

    End If

    If Not IsNull(can_use_restricted) Then
        params.Add "can_use_restricted", can_use_restricted

    End If

    If Not IsNull(min_payment_amount) Then
        params.Add "min_payment_amount", min_payment_amount

    End If

    Set res = makeRequest("SetChildAccountInfo", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set SetChildAccountInfo = res
End Function

Public Function GetResourcePrice(Optional resource_type = Null, Optional price_group_id = Null, Optional price_group_name = Null, Optional resource_param = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(resource_type) Then
        params.Add "resource_type", serialize_list(resource_type)

    End If

    If Not IsNull(price_group_id) Then
        params.Add "price_group_id", serialize_list(price_group_id)

    End If

    If Not IsNull(price_group_name) Then
        params.Add "price_group_name", price_group_name

    End If

    If Not IsNull(resource_param) Then
        params.Add "resource_param", serialize_list(resource_param)

    End If

    Set res = makeRequest("GetResourcePrice", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetResourcePrice = res
End Function

Public Function GetSubscriptionPrice(Optional subscription_template_id = Null, Optional subscription_template_type = Null, Optional subscription_template_name = Null, Optional count = Null, Optional offset = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(subscription_template_id) Then
        params.Add "subscription_template_id", serialize_list(subscription_template_id)

    End If

    If Not IsNull(subscription_template_type) Then
        params.Add "subscription_template_type", subscription_template_type

    End If

    If Not IsNull(subscription_template_name) Then
        params.Add "subscription_template_name", subscription_template_name

    End If

    If Not IsNull(count) Then
        params.Add "count", count

    End If

    If Not IsNull(offset) Then
        params.Add "offset", offset

    End If

    Set res = makeRequest("GetSubscriptionPrice", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetSubscriptionPrice = res
End Function

Public Function GetChildrenAccounts(Optional child_account_id = Null, Optional child_account_name = Null, Optional child_account_email = Null, Optional active = Null, Optional frozen = Null, Optional ignore_invalid_accounts = Null, Optional brief_output = Null, Optional medium_output = Null, Optional count = Null, Optional offset = Null, Optional order_by = Null, Optional return_live_balance = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(child_account_id) Then
        params.Add "child_account_id", serialize_list(child_account_id)

    End If

    If Not IsNull(child_account_name) Then
        params.Add "child_account_name", child_account_name

    End If

    If Not IsNull(child_account_email) Then
        params.Add "child_account_email", child_account_email

    End If

    If Not IsNull(active) Then
        params.Add "active", active

    End If

    If Not IsNull(frozen) Then
        params.Add "frozen", frozen

    End If

    If Not IsNull(ignore_invalid_accounts) Then
        params.Add "ignore_invalid_accounts", ignore_invalid_accounts

    End If

    If Not IsNull(brief_output) Then
        params.Add "brief_output", brief_output

    End If

    If Not IsNull(medium_output) Then
        params.Add "medium_output", medium_output

    End If

    If Not IsNull(count) Then
        params.Add "count", count

    End If

    If Not IsNull(offset) Then
        params.Add "offset", offset

    End If

    If Not IsNull(order_by) Then
        params.Add "order_by", order_by

    End If

    If Not IsNull(return_live_balance) Then
        params.Add "return_live_balance", return_live_balance

    End If

    Set res = makeRequest("GetChildrenAccounts", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetChildrenAccounts = res
End Function

Public Function ChargeAccount(Optional phone_id = Null, Optional phone_number = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(phone_id) Then
        params.Add "phone_id", serialize_list(phone_id)

    End If

    If Not IsNull(phone_number) Then
        params.Add "phone_number", serialize_list(phone_number)

    End If

    Set res = makeRequest("ChargeAccount", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set ChargeAccount = res
End Function

Public Function AddApplication(application_name, Optional secure_record_storage = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "application_name", application_name



    If Not IsNull(secure_record_storage) Then
        params.Add "secure_record_storage", secure_record_storage

    End If

    Set res = makeRequest("AddApplication", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set AddApplication = res
End Function

Public Function DelApplication(Optional application_id = Null, Optional application_name = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(application_id) Then
        params.Add "application_id", serialize_list(application_id)

    End If

    If Not IsNull(application_name) Then
        params.Add "application_name", serialize_list(application_name)

    End If

    Set res = makeRequest("DelApplication", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set DelApplication = res
End Function

Public Function SetApplicationInfo(Optional application_id = Null, Optional required_application_name = Null, Optional application_name = Null, Optional secure_record_storage = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(application_id) Then
        params.Add "application_id", application_id

    End If

    If Not IsNull(required_application_name) Then
        params.Add "required_application_name", required_application_name

    End If

    If Not IsNull(application_name) Then
        params.Add "application_name", application_name

    End If

    If Not IsNull(secure_record_storage) Then
        params.Add "secure_record_storage", secure_record_storage

    End If

    Set res = makeRequest("SetApplicationInfo", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set SetApplicationInfo = res
End Function

Public Function GetApplications(Optional application_id = Null, Optional application_name = Null, Optional user_id = Null, Optional excluded_user_id = Null, Optional showing_user_id = Null, Optional with_rules = Null, Optional with_scenarios = Null, Optional count = Null, Optional offset = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(application_id) Then
        params.Add "application_id", application_id

    End If

    If Not IsNull(application_name) Then
        params.Add "application_name", application_name

    End If

    If Not IsNull(user_id) Then
        params.Add "user_id", user_id

    End If

    If Not IsNull(excluded_user_id) Then
        params.Add "excluded_user_id", excluded_user_id

    End If

    If Not IsNull(showing_user_id) Then
        params.Add "showing_user_id", showing_user_id

    End If

    If Not IsNull(with_rules) Then
        params.Add "with_rules", with_rules

    End If

    If Not IsNull(with_scenarios) Then
        params.Add "with_scenarios", with_scenarios

    End If

    If Not IsNull(count) Then
        params.Add "count", count

    End If

    If Not IsNull(offset) Then
        params.Add "offset", offset

    End If

    Set res = makeRequest("GetApplications", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetApplications = res
End Function

Public Function AddUser(user_name, user_display_name, user_password, Optional application_id = Null, Optional application_name = Null, Optional parent_accounting = Null, Optional user_active = Null, Optional user_custom_data = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "user_name", user_name

    params.Add "user_display_name", user_display_name

    params.Add "user_password", user_password



    If Not IsNull(application_id) Then
        params.Add "application_id", application_id

    End If

    If Not IsNull(application_name) Then
        params.Add "application_name", application_name

    End If

    If Not IsNull(parent_accounting) Then
        params.Add "parent_accounting", parent_accounting

    End If

    If Not IsNull(user_active) Then
        params.Add "user_active", user_active

    End If

    If Not IsNull(user_custom_data) Then
        params.Add "user_custom_data", user_custom_data

    End If

    Set res = makeRequest("AddUser", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set AddUser = res
End Function

Public Function DelUser(Optional user_id = Null, Optional user_name = Null, Optional application_id = Null, Optional application_name = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(user_id) Then
        params.Add "user_id", serialize_list(user_id)

    End If

    If Not IsNull(user_name) Then
        params.Add "user_name", serialize_list(user_name)

    End If

    If Not IsNull(application_id) Then
        params.Add "application_id", application_id

    End If

    If Not IsNull(application_name) Then
        params.Add "application_name", application_name

    End If

    Set res = makeRequest("DelUser", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set DelUser = res
End Function

Public Function SetUserInfo(Optional user_id = Null, Optional user_name = Null, Optional application_id = Null, Optional application_name = Null, Optional new_user_name = Null, Optional user_display_name = Null, Optional user_password = Null, Optional parent_accounting = Null, Optional user_active = Null, Optional user_custom_data = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(user_id) Then
        params.Add "user_id", user_id

    End If

    If Not IsNull(user_name) Then
        params.Add "user_name", user_name

    End If

    If Not IsNull(application_id) Then
        params.Add "application_id", application_id

    End If

    If Not IsNull(application_name) Then
        params.Add "application_name", application_name

    End If

    If Not IsNull(new_user_name) Then
        params.Add "new_user_name", new_user_name

    End If

    If Not IsNull(user_display_name) Then
        params.Add "user_display_name", user_display_name

    End If

    If Not IsNull(user_password) Then
        params.Add "user_password", user_password

    End If

    If Not IsNull(parent_accounting) Then
        params.Add "parent_accounting", parent_accounting

    End If

    If Not IsNull(user_active) Then
        params.Add "user_active", user_active

    End If

    If Not IsNull(user_custom_data) Then
        params.Add "user_custom_data", user_custom_data

    End If

    Set res = makeRequest("SetUserInfo", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set SetUserInfo = res
End Function

Public Function GetUsers(Optional application_id = Null, Optional application_name = Null, Optional skill_id = Null, Optional excluded_skill_id = Null, Optional acd_queue_id = Null, Optional excluded_acd_queue_id = Null, Optional user_id = Null, Optional user_name = Null, Optional user_active = Null, Optional user_display_name = Null, Optional with_skills = Null, Optional with_queues = Null, Optional acd_status = Null, Optional showing_skill_id = Null, Optional count = Null, Optional offset = Null, Optional order_by = Null, Optional return_live_balance = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(application_id) Then
        params.Add "application_id", application_id

    End If

    If Not IsNull(application_name) Then
        params.Add "application_name", application_name

    End If

    If Not IsNull(skill_id) Then
        params.Add "skill_id", skill_id

    End If

    If Not IsNull(excluded_skill_id) Then
        params.Add "excluded_skill_id", excluded_skill_id

    End If

    If Not IsNull(acd_queue_id) Then
        params.Add "acd_queue_id", acd_queue_id

    End If

    If Not IsNull(excluded_acd_queue_id) Then
        params.Add "excluded_acd_queue_id", excluded_acd_queue_id

    End If

    If Not IsNull(user_id) Then
        params.Add "user_id", user_id

    End If

    If Not IsNull(user_name) Then
        params.Add "user_name", user_name

    End If

    If Not IsNull(user_active) Then
        params.Add "user_active", user_active

    End If

    If Not IsNull(user_display_name) Then
        params.Add "user_display_name", user_display_name

    End If

    If Not IsNull(with_skills) Then
        params.Add "with_skills", with_skills

    End If

    If Not IsNull(with_queues) Then
        params.Add "with_queues", with_queues

    End If

    If Not IsNull(acd_status) Then
        params.Add "acd_status", serialize_list(acd_status)

    End If

    If Not IsNull(showing_skill_id) Then
        params.Add "showing_skill_id", showing_skill_id

    End If

    If Not IsNull(count) Then
        params.Add "count", count

    End If

    If Not IsNull(offset) Then
        params.Add "offset", offset

    End If

    If Not IsNull(order_by) Then
        params.Add "order_by", order_by

    End If

    If Not IsNull(return_live_balance) Then
        params.Add "return_live_balance", return_live_balance

    End If

    Set res = makeRequest("GetUsers", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetUsers = res
End Function

Public Function CreateCallList(rule_id, priority, max_simultaneous, num_attempts, name, file_content, Optional interval_seconds = Null, Optional queue_id = Null, Optional avg_waiting_sec = Null, Optional encoding = Null, Optional delimiter = Null, Optional escape = Null, Optional reference_ip = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "rule_id", rule_id

    params.Add "priority", priority

    params.Add "max_simultaneous", max_simultaneous

    params.Add "num_attempts", num_attempts

    params.Add "name", name

    params.Add "file_content", file_content



    If Not IsNull(interval_seconds) Then
        params.Add "interval_seconds", interval_seconds

    End If

    If Not IsNull(queue_id) Then
        params.Add "queue_id", queue_id

    End If

    If Not IsNull(avg_waiting_sec) Then
        params.Add "avg_waiting_sec", avg_waiting_sec

    End If

    If Not IsNull(encoding) Then
        params.Add "encoding", encoding

    End If

    If Not IsNull(delimiter) Then
        params.Add "delimiter", delimiter

    End If

    If Not IsNull(escape) Then
        params.Add "escape", escape

    End If

    If Not IsNull(reference_ip) Then
        params.Add "reference_ip", reference_ip

    End If

    Set res = makeRequest("CreateCallList", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set CreateCallList = res
End Function

Public Function CreateManualCallList(rule_id, priority, max_simultaneous, num_attempts, name, file_content, Optional interval_seconds = Null, Optional encoding = Null, Optional delimiter = Null, Optional escape = Null, Optional reference_ip = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "rule_id", rule_id

    params.Add "priority", priority

    params.Add "max_simultaneous", max_simultaneous

    params.Add "num_attempts", num_attempts

    params.Add "name", name

    params.Add "file_content", file_content



    If Not IsNull(interval_seconds) Then
        params.Add "interval_seconds", interval_seconds

    End If

    If Not IsNull(encoding) Then
        params.Add "encoding", encoding

    End If

    If Not IsNull(delimiter) Then
        params.Add "delimiter", delimiter

    End If

    If Not IsNull(escape) Then
        params.Add "escape", escape

    End If

    If Not IsNull(reference_ip) Then
        params.Add "reference_ip", reference_ip

    End If

    Set res = makeRequest("CreateManualCallList", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set CreateManualCallList = res
End Function

Public Function StartNextCallTask(list_id, Optional custom_params = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "list_id", serialize_list(list_id)



    If Not IsNull(custom_params) Then
        params.Add "custom_params", custom_params

    End If

    Set res = makeRequest("StartNextCallTask", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set StartNextCallTask = res
End Function

Public Function AppendToCallList(file_content, Optional list_id = Null, Optional list_name = Null, Optional encoding = Null, Optional escape = Null, Optional delimiter = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "file_content", file_content



    If Not IsNull(list_id) Then
        params.Add "list_id", list_id

    End If

    If Not IsNull(list_name) Then
        params.Add "list_name", list_name

    End If

    If Not IsNull(encoding) Then
        params.Add "encoding", encoding

    End If

    If Not IsNull(escape) Then
        params.Add "escape", escape

    End If

    If Not IsNull(delimiter) Then
        params.Add "delimiter", delimiter

    End If

    Set res = makeRequest("AppendToCallList", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set AppendToCallList = res
End Function

Public Function GetCallLists(Optional list_id = Null, Optional name = Null, Optional is_active = Null, Optional from_date = Null, Optional to_date = Null, Optional type_list = Null, Optional count = Null, Optional offset = Null, Optional application_id = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(list_id) Then
        params.Add "list_id", serialize_list(list_id)

    End If

    If Not IsNull(name) Then
        params.Add "name", name

    End If

    If Not IsNull(is_active) Then
        params.Add "is_active", is_active

    End If

    If Not IsNull(from_date) Then
        params.Add "from_date", vba_datetime_to_api(from_date)

    End If

    If Not IsNull(to_date) Then
        params.Add "to_date", vba_datetime_to_api(to_date)

    End If

    If Not IsNull(type_list) Then
        params.Add "type_list", type_list

    End If

    If Not IsNull(count) Then
        params.Add "count", count

    End If

    If Not IsNull(offset) Then
        params.Add "offset", offset

    End If

    If Not IsNull(application_id) Then
        params.Add "application_id", serialize_list(application_id)

    End If

    Set res = makeRequest("GetCallLists", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetCallLists = res
End Function

Public Function GetCallListDetails(list_id, Optional count = Null, Optional offset = Null, Optional output = Null, Optional encoding = Null, Optional delimiter = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "list_id", list_id



    If Not IsNull(count) Then
        params.Add "count", count

    End If

    If Not IsNull(offset) Then
        params.Add "offset", offset

    End If

    If Not IsNull(output) Then
        params.Add "output", output

    End If

    If Not IsNull(encoding) Then
        params.Add "encoding", encoding

    End If

    If Not IsNull(delimiter) Then
        params.Add "delimiter", delimiter

    End If

    Set res = makeRequest("GetCallListDetails", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetCallListDetails = res
End Function

Public Function StopCallListProcessing(list_id) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "list_id", list_id



    Set res = makeRequest("StopCallListProcessing", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set StopCallListProcessing = res
End Function

Public Function RecoverCallList(list_id) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "list_id", list_id



    Set res = makeRequest("RecoverCallList", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set RecoverCallList = res
End Function

Public Function AddScenario(scenario_name, Optional scenario_script = Null, Optional rule_id = Null, Optional rule_name = Null, Optional rewrite = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "scenario_name", scenario_name



    If Not IsNull(scenario_script) Then
        params.Add "scenario_script", scenario_script

    End If

    If Not IsNull(rule_id) Then
        params.Add "rule_id", rule_id

    End If

    If Not IsNull(rule_name) Then
        params.Add "rule_name", rule_name

    End If

    If Not IsNull(rewrite) Then
        params.Add "rewrite", rewrite

    End If

    Set res = makeRequest("AddScenario", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set AddScenario = res
End Function

Public Function DelScenario(Optional scenario_id = Null, Optional scenario_name = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(scenario_id) Then
        params.Add "scenario_id", serialize_list(scenario_id)

    End If

    If Not IsNull(scenario_name) Then
        params.Add "scenario_name", serialize_list(scenario_name)

    End If

    Set res = makeRequest("DelScenario", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set DelScenario = res
End Function

Public Function BindScenario(Optional scenario_id = Null, Optional scenario_name = Null, Optional rule_id = Null, Optional rule_name = Null, Optional application_id = Null, Optional application_name = Null, Optional bind = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(scenario_id) Then
        params.Add "scenario_id", serialize_list(scenario_id)

    End If

    If Not IsNull(scenario_name) Then
        params.Add "scenario_name", serialize_list(scenario_name)

    End If

    If Not IsNull(rule_id) Then
        params.Add "rule_id", rule_id

    End If

    If Not IsNull(rule_name) Then
        params.Add "rule_name", rule_name

    End If

    If Not IsNull(application_id) Then
        params.Add "application_id", application_id

    End If

    If Not IsNull(application_name) Then
        params.Add "application_name", application_name

    End If

    If Not IsNull(bind) Then
        params.Add "bind", bind

    End If

    Set res = makeRequest("BindScenario", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set BindScenario = res
End Function

Public Function GetScenarios(Optional scenario_id = Null, Optional scenario_name = Null, Optional with_script = Null, Optional count = Null, Optional offset = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(scenario_id) Then
        params.Add "scenario_id", scenario_id

    End If

    If Not IsNull(scenario_name) Then
        params.Add "scenario_name", scenario_name

    End If

    If Not IsNull(with_script) Then
        params.Add "with_script", with_script

    End If

    If Not IsNull(count) Then
        params.Add "count", count

    End If

    If Not IsNull(offset) Then
        params.Add "offset", offset

    End If

    Set res = makeRequest("GetScenarios", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetScenarios = res
End Function

Public Function SetScenarioInfo(Optional scenario_id = Null, Optional required_scenario_name = Null, Optional scenario_name = Null, Optional scenario_script = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(scenario_id) Then
        params.Add "scenario_id", scenario_id

    End If

    If Not IsNull(required_scenario_name) Then
        params.Add "required_scenario_name", required_scenario_name

    End If

    If Not IsNull(scenario_name) Then
        params.Add "scenario_name", scenario_name

    End If

    If Not IsNull(scenario_script) Then
        params.Add "scenario_script", scenario_script

    End If

    Set res = makeRequest("SetScenarioInfo", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set SetScenarioInfo = res
End Function

Public Function ReorderScenarios(Optional rule_id = Null, Optional rule_name = Null, Optional scenario_id = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(rule_id) Then
        params.Add "rule_id", rule_id

    End If

    If Not IsNull(rule_name) Then
        params.Add "rule_name", rule_name

    End If

    If Not IsNull(scenario_id) Then
        params.Add "scenario_id", serialize_list(scenario_id)

    End If

    Set res = makeRequest("ReorderScenarios", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set ReorderScenarios = res
End Function

Public Function StartScenarios(rule_id, Optional user_id = Null, Optional user_name = Null, Optional application_id = Null, Optional application_name = Null, Optional script_custom_data = Null, Optional reference_ip = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "rule_id", rule_id



    If Not IsNull(user_id) Then
        params.Add "user_id", user_id

    End If

    If Not IsNull(user_name) Then
        params.Add "user_name", user_name

    End If

    If Not IsNull(application_id) Then
        params.Add "application_id", application_id

    End If

    If Not IsNull(application_name) Then
        params.Add "application_name", application_name

    End If

    If Not IsNull(script_custom_data) Then
        params.Add "script_custom_data", script_custom_data

    End If

    If Not IsNull(reference_ip) Then
        params.Add "reference_ip", reference_ip

    End If

    Set res = makeRequest("StartScenarios", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set StartScenarios = res
End Function

Public Function StartConference(conference_name, rule_id, Optional user_id = Null, Optional user_name = Null, Optional application_id = Null, Optional application_name = Null, Optional script_custom_data = Null, Optional reference_ip = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "conference_name", conference_name

    params.Add "rule_id", rule_id



    If Not IsNull(user_id) Then
        params.Add "user_id", user_id

    End If

    If Not IsNull(user_name) Then
        params.Add "user_name", user_name

    End If

    If Not IsNull(application_id) Then
        params.Add "application_id", application_id

    End If

    If Not IsNull(application_name) Then
        params.Add "application_name", application_name

    End If

    If Not IsNull(script_custom_data) Then
        params.Add "script_custom_data", script_custom_data

    End If

    If Not IsNull(reference_ip) Then
        params.Add "reference_ip", reference_ip

    End If

    Set res = makeRequest("StartConference", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set StartConference = res
End Function

Public Function AddRule(rule_name, rule_pattern, Optional application_id = Null, Optional application_name = Null, Optional rule_pattern_exclude = Null, Optional video_conference = Null, Optional scenario_id = Null, Optional scenario_name = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "rule_name", rule_name

    params.Add "rule_pattern", rule_pattern



    If Not IsNull(application_id) Then
        params.Add "application_id", application_id

    End If

    If Not IsNull(application_name) Then
        params.Add "application_name", application_name

    End If

    If Not IsNull(rule_pattern_exclude) Then
        params.Add "rule_pattern_exclude", rule_pattern_exclude

    End If

    If Not IsNull(video_conference) Then
        params.Add "video_conference", video_conference

    End If

    If Not IsNull(scenario_id) Then
        params.Add "scenario_id", serialize_list(scenario_id)

    End If

    If Not IsNull(scenario_name) Then
        params.Add "scenario_name", serialize_list(scenario_name)

    End If

    Set res = makeRequest("AddRule", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set AddRule = res
End Function

Public Function DelRule(Optional rule_id = Null, Optional rule_name = Null, Optional application_id = Null, Optional application_name = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(rule_id) Then
        params.Add "rule_id", serialize_list(rule_id)

    End If

    If Not IsNull(rule_name) Then
        params.Add "rule_name", serialize_list(rule_name)

    End If

    If Not IsNull(application_id) Then
        params.Add "application_id", serialize_list(application_id)

    End If

    If Not IsNull(application_name) Then
        params.Add "application_name", serialize_list(application_name)

    End If

    Set res = makeRequest("DelRule", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set DelRule = res
End Function

Public Function SetRuleInfo(rule_id, Optional rule_name = Null, Optional rule_pattern = Null, Optional rule_pattern_exclude = Null, Optional video_conference = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "rule_id", rule_id



    If Not IsNull(rule_name) Then
        params.Add "rule_name", rule_name

    End If

    If Not IsNull(rule_pattern) Then
        params.Add "rule_pattern", rule_pattern

    End If

    If Not IsNull(rule_pattern_exclude) Then
        params.Add "rule_pattern_exclude", rule_pattern_exclude

    End If

    If Not IsNull(video_conference) Then
        params.Add "video_conference", video_conference

    End If

    Set res = makeRequest("SetRuleInfo", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set SetRuleInfo = res
End Function

Public Function GetRules(Optional application_id = Null, Optional application_name = Null, Optional rule_id = Null, Optional rule_name = Null, Optional video_conference = Null, Optional template = Null, Optional with_scenarios = Null, Optional count = Null, Optional offset = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(application_id) Then
        params.Add "application_id", application_id

    End If

    If Not IsNull(application_name) Then
        params.Add "application_name", application_name

    End If

    If Not IsNull(rule_id) Then
        params.Add "rule_id", rule_id

    End If

    If Not IsNull(rule_name) Then
        params.Add "rule_name", rule_name

    End If

    If Not IsNull(video_conference) Then
        params.Add "video_conference", video_conference

    End If

    If Not IsNull(template) Then
        params.Add "template", template

    End If

    If Not IsNull(with_scenarios) Then
        params.Add "with_scenarios", with_scenarios

    End If

    If Not IsNull(count) Then
        params.Add "count", count

    End If

    If Not IsNull(offset) Then
        params.Add "offset", offset

    End If

    Set res = makeRequest("GetRules", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetRules = res
End Function

Public Function ReorderRules(rule_id) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "rule_id", serialize_list(rule_id)



    Set res = makeRequest("ReorderRules", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set ReorderRules = res
End Function

Public Function GetCallHistory(from_date, to_date, Optional call_session_history_id = Null, Optional application_id = Null, Optional application_name = Null, Optional user_id = Null, Optional rule_name = Null, Optional remote_number = Null, Optional local_number = Null, Optional call_session_history_custom_data = Null, Optional with_calls = Null, Optional with_records = Null, Optional with_other_resources = Null, Optional child_account_id = Null, Optional children_calls_only = Null, Optional with_header = Null, Optional desc_order = Null, Optional with_total_count = Null, Optional count = Null, Optional offset = Null, Optional output = Null, Optional is_async = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "from_date", vba_datetime_to_api(from_date)

    params.Add "to_date", vba_datetime_to_api(to_date)



    If Not IsNull(call_session_history_id) Then
        params.Add "call_session_history_id", serialize_list(call_session_history_id)

    End If

    If Not IsNull(application_id) Then
        params.Add "application_id", application_id

    End If

    If Not IsNull(application_name) Then
        params.Add "application_name", application_name

    End If

    If Not IsNull(user_id) Then
        params.Add "user_id", serialize_list(user_id)

    End If

    If Not IsNull(rule_name) Then
        params.Add "rule_name", rule_name

    End If

    If Not IsNull(remote_number) Then
        params.Add "remote_number", serialize_list(remote_number)

    End If

    If Not IsNull(local_number) Then
        params.Add "local_number", serialize_list(local_number)

    End If

    If Not IsNull(call_session_history_custom_data) Then
        params.Add "call_session_history_custom_data", call_session_history_custom_data

    End If

    If Not IsNull(with_calls) Then
        params.Add "with_calls", with_calls

    End If

    If Not IsNull(with_records) Then
        params.Add "with_records", with_records

    End If

    If Not IsNull(with_other_resources) Then
        params.Add "with_other_resources", with_other_resources

    End If

    If Not IsNull(child_account_id) Then
        params.Add "child_account_id", serialize_list(child_account_id)

    End If

    If Not IsNull(children_calls_only) Then
        params.Add "children_calls_only", children_calls_only

    End If

    If Not IsNull(with_header) Then
        params.Add "with_header", with_header

    End If

    If Not IsNull(desc_order) Then
        params.Add "desc_order", desc_order

    End If

    If Not IsNull(with_total_count) Then
        params.Add "with_total_count", with_total_count

    End If

    If Not IsNull(count) Then
        params.Add "count", count

    End If

    If Not IsNull(offset) Then
        params.Add "offset", offset

    End If

    If Not IsNull(output) Then
        params.Add "output", output

    End If

    If Not IsNull(is_async) Then
        params.Add "is_async", is_async

    End If

    Set res = makeRequest("GetCallHistory", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetCallHistory = res
End Function

Public Function GetHistoryReports(Optional history_report_id = Null, Optional history_type = Null, Optional created_from = Null, Optional created_to = Null, Optional is_completed = Null, Optional desc_order = Null, Optional count = Null, Optional offset = Null, Optional application_id = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(history_report_id) Then
        params.Add "history_report_id", history_report_id

    End If

    If Not IsNull(history_type) Then
        params.Add "history_type", serialize_list(history_type)

    End If

    If Not IsNull(created_from) Then
        params.Add "created_from", vba_datetime_to_api(created_from)

    End If

    If Not IsNull(created_to) Then
        params.Add "created_to", vba_datetime_to_api(created_to)

    End If

    If Not IsNull(is_completed) Then
        params.Add "is_completed", is_completed

    End If

    If Not IsNull(desc_order) Then
        params.Add "desc_order", desc_order

    End If

    If Not IsNull(count) Then
        params.Add "count", count

    End If

    If Not IsNull(offset) Then
        params.Add "offset", offset

    End If

    If Not IsNull(application_id) Then
        params.Add "application_id", serialize_list(application_id)

    End If

    Set res = makeRequest("GetHistoryReports", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetHistoryReports = res
End Function

Public Function GetTransactionHistory(from_date, to_date, Optional transaction_id = Null, Optional transaction_type = Null, Optional user_id = Null, Optional child_account_id = Null, Optional children_transactions_only = Null, Optional users_transactions_only = Null, Optional desc_order = Null, Optional count = Null, Optional offset = Null, Optional output = Null, Optional is_async = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "from_date", vba_datetime_to_api(from_date)

    params.Add "to_date", vba_datetime_to_api(to_date)



    If Not IsNull(transaction_id) Then
        params.Add "transaction_id", serialize_list(transaction_id)

    End If

    If Not IsNull(transaction_type) Then
        params.Add "transaction_type", serialize_list(transaction_type)

    End If

    If Not IsNull(user_id) Then
        params.Add "user_id", serialize_list(user_id)

    End If

    If Not IsNull(child_account_id) Then
        params.Add "child_account_id", serialize_list(child_account_id)

    End If

    If Not IsNull(children_transactions_only) Then
        params.Add "children_transactions_only", children_transactions_only

    End If

    If Not IsNull(users_transactions_only) Then
        params.Add "users_transactions_only", users_transactions_only

    End If

    If Not IsNull(desc_order) Then
        params.Add "desc_order", desc_order

    End If

    If Not IsNull(count) Then
        params.Add "count", count

    End If

    If Not IsNull(offset) Then
        params.Add "offset", offset

    End If

    If Not IsNull(output) Then
        params.Add "output", output

    End If

    If Not IsNull(is_async) Then
        params.Add "is_async", is_async

    End If

    Set res = makeRequest("GetTransactionHistory", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetTransactionHistory = res
End Function

Public Function DeleteRecord(Optional record_url = Null, Optional record_id = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(record_url) Then
        params.Add "record_url", record_url

    End If

    If Not IsNull(record_id) Then
        params.Add "record_id", record_id

    End If

    Set res = makeRequest("DeleteRecord", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set DeleteRecord = res
End Function

Public Function GetACDHistory(from_date, to_date, Optional acd_session_history_id = Null, Optional acd_request_id = Null, Optional acd_queue_id = Null, Optional user_id = Null, Optional operator_hangup = Null, Optional unserviced = Null, Optional min_waiting_time = Null, Optional rejected = Null, Optional with_events = Null, Optional with_header = Null, Optional desc_order = Null, Optional count = Null, Optional offset = Null, Optional output = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "from_date", vba_datetime_to_api(from_date)

    params.Add "to_date", vba_datetime_to_api(to_date)



    If Not IsNull(acd_session_history_id) Then
        params.Add "acd_session_history_id", serialize_list(acd_session_history_id)

    End If

    If Not IsNull(acd_request_id) Then
        params.Add "acd_request_id", serialize_list(acd_request_id)

    End If

    If Not IsNull(acd_queue_id) Then
        params.Add "acd_queue_id", serialize_list(acd_queue_id)

    End If

    If Not IsNull(user_id) Then
        params.Add "user_id", serialize_list(user_id)

    End If

    If Not IsNull(operator_hangup) Then
        params.Add "operator_hangup", operator_hangup

    End If

    If Not IsNull(unserviced) Then
        params.Add "unserviced", unserviced

    End If

    If Not IsNull(min_waiting_time) Then
        params.Add "min_waiting_time", min_waiting_time

    End If

    If Not IsNull(rejected) Then
        params.Add "rejected", rejected

    End If

    If Not IsNull(with_events) Then
        params.Add "with_events", with_events

    End If

    If Not IsNull(with_header) Then
        params.Add "with_header", with_header

    End If

    If Not IsNull(desc_order) Then
        params.Add "desc_order", desc_order

    End If

    If Not IsNull(count) Then
        params.Add "count", count

    End If

    If Not IsNull(offset) Then
        params.Add "offset", offset

    End If

    If Not IsNull(output) Then
        params.Add "output", output

    End If

    Set res = makeRequest("GetACDHistory", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetACDHistory = res
End Function

Public Function GetAuditLog(from_date, to_date, Optional audit_log_id = Null, Optional filtered_admin_user_id = Null, Optional filtered_ip = Null, Optional filtered_cmd = Null, Optional advanced_filters = Null, Optional with_header = Null, Optional desc_order = Null, Optional with_total_count = Null, Optional count = Null, Optional offset = Null, Optional output = Null, Optional is_async = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "from_date", vba_datetime_to_api(from_date)

    params.Add "to_date", vba_datetime_to_api(to_date)



    If Not IsNull(audit_log_id) Then
        params.Add "audit_log_id", serialize_list(audit_log_id)

    End If

    If Not IsNull(filtered_admin_user_id) Then
        params.Add "filtered_admin_user_id", filtered_admin_user_id

    End If

    If Not IsNull(filtered_ip) Then
        params.Add "filtered_ip", serialize_list(filtered_ip)

    End If

    If Not IsNull(filtered_cmd) Then
        params.Add "filtered_cmd", serialize_list(filtered_cmd)

    End If

    If Not IsNull(advanced_filters) Then
        params.Add "advanced_filters", advanced_filters

    End If

    If Not IsNull(with_header) Then
        params.Add "with_header", with_header

    End If

    If Not IsNull(desc_order) Then
        params.Add "desc_order", desc_order

    End If

    If Not IsNull(with_total_count) Then
        params.Add "with_total_count", with_total_count

    End If

    If Not IsNull(count) Then
        params.Add "count", count

    End If

    If Not IsNull(offset) Then
        params.Add "offset", offset

    End If

    If Not IsNull(output) Then
        params.Add "output", output

    End If

    If Not IsNull(is_async) Then
        params.Add "is_async", is_async

    End If

    Set res = makeRequest("GetAuditLog", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetAuditLog = res
End Function

Public Function AddPstnBlackListItem(pstn_blacklist_phone) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "pstn_blacklist_phone", pstn_blacklist_phone



    Set res = makeRequest("AddPstnBlackListItem", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set AddPstnBlackListItem = res
End Function

Public Function SetPstnBlackListItem(pstn_blacklist_id, pstn_blacklist_phone) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "pstn_blacklist_id", pstn_blacklist_id

    params.Add "pstn_blacklist_phone", pstn_blacklist_phone



    Set res = makeRequest("SetPstnBlackListItem", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set SetPstnBlackListItem = res
End Function

Public Function DelPstnBlackListItem(pstn_blacklist_id) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "pstn_blacklist_id", pstn_blacklist_id



    Set res = makeRequest("DelPstnBlackListItem", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set DelPstnBlackListItem = res
End Function

Public Function GetPstnBlackList(Optional pstn_blacklist_id = Null, Optional pstn_blacklist_phone = Null, Optional count = Null, Optional offset = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(pstn_blacklist_id) Then
        params.Add "pstn_blacklist_id", pstn_blacklist_id

    End If

    If Not IsNull(pstn_blacklist_phone) Then
        params.Add "pstn_blacklist_phone", pstn_blacklist_phone

    End If

    If Not IsNull(count) Then
        params.Add "count", count

    End If

    If Not IsNull(offset) Then
        params.Add "offset", offset

    End If

    Set res = makeRequest("GetPstnBlackList", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetPstnBlackList = res
End Function

Public Function AddSipWhiteListItem(sip_whitelist_network) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "sip_whitelist_network", sip_whitelist_network



    Set res = makeRequest("AddSipWhiteListItem", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set AddSipWhiteListItem = res
End Function

Public Function DelSipWhiteListItem(sip_whitelist_id) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "sip_whitelist_id", sip_whitelist_id



    Set res = makeRequest("DelSipWhiteListItem", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set DelSipWhiteListItem = res
End Function

Public Function SetSipWhiteListItem(sip_whitelist_id, sip_whitelist_network) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "sip_whitelist_id", sip_whitelist_id

    params.Add "sip_whitelist_network", sip_whitelist_network



    Set res = makeRequest("SetSipWhiteListItem", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set SetSipWhiteListItem = res
End Function

Public Function GetSipWhiteList(Optional sip_whitelist_id = Null, Optional count = Null, Optional offset = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(sip_whitelist_id) Then
        params.Add "sip_whitelist_id", sip_whitelist_id

    End If

    If Not IsNull(count) Then
        params.Add "count", count

    End If

    If Not IsNull(offset) Then
        params.Add "offset", offset

    End If

    Set res = makeRequest("GetSipWhiteList", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetSipWhiteList = res
End Function

Public Function CreateSipRegistration(sip_username, proxy, Optional auth_user = Null, Optional outbound_proxy = Null, Optional password = Null, Optional is_persistent = Null, Optional application_id = Null, Optional application_name = Null, Optional rule_id = Null, Optional rule_name = Null, Optional user_id = Null, Optional user_name = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "sip_username", sip_username

    params.Add "proxy", proxy



    If Not IsNull(auth_user) Then
        params.Add "auth_user", auth_user

    End If

    If Not IsNull(outbound_proxy) Then
        params.Add "outbound_proxy", outbound_proxy

    End If

    If Not IsNull(password) Then
        params.Add "password", password

    End If

    If Not IsNull(is_persistent) Then
        params.Add "is_persistent", is_persistent

    End If

    If Not IsNull(application_id) Then
        params.Add "application_id", application_id

    End If

    If Not IsNull(application_name) Then
        params.Add "application_name", application_name

    End If

    If Not IsNull(rule_id) Then
        params.Add "rule_id", rule_id

    End If

    If Not IsNull(rule_name) Then
        params.Add "rule_name", rule_name

    End If

    If Not IsNull(user_id) Then
        params.Add "user_id", user_id

    End If

    If Not IsNull(user_name) Then
        params.Add "user_name", user_name

    End If

    Set res = makeRequest("CreateSipRegistration", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set CreateSipRegistration = res
End Function

Public Function UpdateSipRegistration(sip_registration_id, Optional sip_username = Null, Optional proxy = Null, Optional auth_user = Null, Optional outbound_proxy = Null, Optional password = Null, Optional application_id = Null, Optional application_name = Null, Optional rule_id = Null, Optional rule_name = Null, Optional user_id = Null, Optional user_name = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "sip_registration_id", sip_registration_id



    If Not IsNull(sip_username) Then
        params.Add "sip_username", sip_username

    End If

    If Not IsNull(proxy) Then
        params.Add "proxy", proxy

    End If

    If Not IsNull(auth_user) Then
        params.Add "auth_user", auth_user

    End If

    If Not IsNull(outbound_proxy) Then
        params.Add "outbound_proxy", outbound_proxy

    End If

    If Not IsNull(password) Then
        params.Add "password", password

    End If

    If Not IsNull(application_id) Then
        params.Add "application_id", application_id

    End If

    If Not IsNull(application_name) Then
        params.Add "application_name", application_name

    End If

    If Not IsNull(rule_id) Then
        params.Add "rule_id", rule_id

    End If

    If Not IsNull(rule_name) Then
        params.Add "rule_name", rule_name

    End If

    If Not IsNull(user_id) Then
        params.Add "user_id", user_id

    End If

    If Not IsNull(user_name) Then
        params.Add "user_name", user_name

    End If

    Set res = makeRequest("UpdateSipRegistration", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set UpdateSipRegistration = res
End Function

Public Function BindSipRegistration(Optional sip_registration_id = Null, Optional application_id = Null, Optional application_name = Null, Optional rule_id = Null, Optional rule_name = Null, Optional user_id = Null, Optional user_name = Null, Optional bind = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(sip_registration_id) Then
        params.Add "sip_registration_id", sip_registration_id

    End If

    If Not IsNull(application_id) Then
        params.Add "application_id", application_id

    End If

    If Not IsNull(application_name) Then
        params.Add "application_name", application_name

    End If

    If Not IsNull(rule_id) Then
        params.Add "rule_id", rule_id

    End If

    If Not IsNull(rule_name) Then
        params.Add "rule_name", rule_name

    End If

    If Not IsNull(user_id) Then
        params.Add "user_id", user_id

    End If

    If Not IsNull(user_name) Then
        params.Add "user_name", user_name

    End If

    If Not IsNull(bind) Then
        params.Add "bind", bind

    End If

    Set res = makeRequest("BindSipRegistration", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set BindSipRegistration = res
End Function

Public Function DeleteSipRegistration(sip_registration_id) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "sip_registration_id", sip_registration_id



    Set res = makeRequest("DeleteSipRegistration", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set DeleteSipRegistration = res
End Function

Public Function GetSipRegistrations(Optional sip_registration_id = Null, Optional sip_username = Null, Optional deactivated = Null, Optional successful = Null, Optional is_persistent = Null, Optional application_id = Null, Optional application_name = Null, Optional is_bound_to_application = Null, Optional rule_id = Null, Optional rule_name = Null, Optional user_id = Null, Optional user_name = Null, Optional proxy = Null, Optional in_progress = Null, Optional status_code = Null, Optional count = Null, Optional offset = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(sip_registration_id) Then
        params.Add "sip_registration_id", sip_registration_id

    End If

    If Not IsNull(sip_username) Then
        params.Add "sip_username", sip_username

    End If

    If Not IsNull(deactivated) Then
        params.Add "deactivated", deactivated

    End If

    If Not IsNull(successful) Then
        params.Add "successful", successful

    End If

    If Not IsNull(is_persistent) Then
        params.Add "is_persistent", is_persistent

    End If

    If Not IsNull(application_id) Then
        params.Add "application_id", serialize_list(application_id)

    End If

    If Not IsNull(application_name) Then
        params.Add "application_name", serialize_list(application_name)

    End If

    If Not IsNull(is_bound_to_application) Then
        params.Add "is_bound_to_application", is_bound_to_application

    End If

    If Not IsNull(rule_id) Then
        params.Add "rule_id", serialize_list(rule_id)

    End If

    If Not IsNull(rule_name) Then
        params.Add "rule_name", serialize_list(rule_name)

    End If

    If Not IsNull(user_id) Then
        params.Add "user_id", serialize_list(user_id)

    End If

    If Not IsNull(user_name) Then
        params.Add "user_name", serialize_list(user_name)

    End If

    If Not IsNull(proxy) Then
        params.Add "proxy", serialize_list(proxy)

    End If

    If Not IsNull(in_progress) Then
        params.Add "in_progress", in_progress

    End If

    If Not IsNull(status_code) Then
        params.Add "status_code", status_code

    End If

    If Not IsNull(count) Then
        params.Add "count", count

    End If

    If Not IsNull(offset) Then
        params.Add "offset", offset

    End If

    Set res = makeRequest("GetSipRegistrations", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetSipRegistrations = res
End Function

Public Function AttachPhoneNumber(country_code, phone_category_name, phone_region_id, Optional phone_count = Null, Optional phone_number = Null, Optional country_state = Null, Optional regulation_address_id = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "country_code", country_code

    params.Add "phone_category_name", phone_category_name

    params.Add "phone_region_id", phone_region_id



    If Not IsNull(phone_count) Then
        params.Add "phone_count", phone_count

    End If

    If Not IsNull(phone_number) Then
        params.Add "phone_number", serialize_list(phone_number)

    End If

    If Not IsNull(country_state) Then
        params.Add "country_state", country_state

    End If

    If Not IsNull(regulation_address_id) Then
        params.Add "regulation_address_id", regulation_address_id

    End If

    Set res = makeRequest("AttachPhoneNumber", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set AttachPhoneNumber = res
End Function

Public Function BindPhoneNumberToApplication(Optional phone_id = Null, Optional phone_number = Null, Optional application_id = Null, Optional application_name = Null, Optional rule_id = Null, Optional rule_name = Null, Optional bind = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(phone_id) Then
        params.Add "phone_id", serialize_list(phone_id)

    End If

    If Not IsNull(phone_number) Then
        params.Add "phone_number", serialize_list(phone_number)

    End If

    If Not IsNull(application_id) Then
        params.Add "application_id", application_id

    End If

    If Not IsNull(application_name) Then
        params.Add "application_name", application_name

    End If

    If Not IsNull(rule_id) Then
        params.Add "rule_id", rule_id

    End If

    If Not IsNull(rule_name) Then
        params.Add "rule_name", rule_name

    End If

    If Not IsNull(bind) Then
        params.Add "bind", bind

    End If

    Set res = makeRequest("BindPhoneNumberToApplication", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set BindPhoneNumberToApplication = res
End Function

Public Function DeactivatePhoneNumber(Optional phone_id = Null, Optional phone_number = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(phone_id) Then
        params.Add "phone_id", serialize_list(phone_id)

    End If

    If Not IsNull(phone_number) Then
        params.Add "phone_number", serialize_list(phone_number)

    End If

    Set res = makeRequest("DeactivatePhoneNumber", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set DeactivatePhoneNumber = res
End Function

Public Function SetPhoneNumberInfo(Optional phone_id = Null, Optional phone_number = Null, Optional incoming_sms_callback_url = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(phone_id) Then
        params.Add "phone_id", serialize_list(phone_id)

    End If

    If Not IsNull(phone_number) Then
        params.Add "phone_number", serialize_list(phone_number)

    End If

    If Not IsNull(incoming_sms_callback_url) Then
        params.Add "incoming_sms_callback_url", incoming_sms_callback_url

    End If

    Set res = makeRequest("SetPhoneNumberInfo", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set SetPhoneNumberInfo = res
End Function

Public Function GetPhoneNumbers(Optional phone_id = Null, Optional application_id = Null, Optional application_name = Null, Optional is_bound_to_application = Null, Optional phone_template = Null, Optional country_code = Null, Optional phone_category_name = Null, Optional canceled = Null, Optional deactivated = Null, Optional auto_charge = Null, Optional from_phone_next_renewal = Null, Optional to_phone_next_renewal = Null, Optional from_phone_purchase_date = Null, Optional to_phone_purchase_date = Null, Optional child_account_id = Null, Optional children_phones_only = Null, Optional verification_name = Null, Optional verification_status = Null, Optional from_unverified_hold_until = Null, Optional to_unverified_hold_until = Null, Optional can_be_used = Null, Optional order_by = Null, Optional sandbox = Null, Optional count = Null, Optional offset = Null, Optional phone_region_name = Null, Optional rule_id = Null, Optional rule_name = Null, Optional is_bound_to_rule = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(phone_id) Then
        params.Add "phone_id", phone_id

    End If

    If Not IsNull(application_id) Then
        params.Add "application_id", application_id

    End If

    If Not IsNull(application_name) Then
        params.Add "application_name", application_name

    End If

    If Not IsNull(is_bound_to_application) Then
        params.Add "is_bound_to_application", is_bound_to_application

    End If

    If Not IsNull(phone_template) Then
        params.Add "phone_template", phone_template

    End If

    If Not IsNull(country_code) Then
        params.Add "country_code", serialize_list(country_code)

    End If

    If Not IsNull(phone_category_name) Then
        params.Add "phone_category_name", phone_category_name

    End If

    If Not IsNull(canceled) Then
        params.Add "canceled", canceled

    End If

    If Not IsNull(deactivated) Then
        params.Add "deactivated", deactivated

    End If

    If Not IsNull(auto_charge) Then
        params.Add "auto_charge", auto_charge

    End If

    If Not IsNull(from_phone_next_renewal) Then
        params.Add "from_phone_next_renewal", from_phone_next_renewal

    End If

    If Not IsNull(to_phone_next_renewal) Then
        params.Add "to_phone_next_renewal", to_phone_next_renewal

    End If

    If Not IsNull(from_phone_purchase_date) Then
        params.Add "from_phone_purchase_date", vba_datetime_to_api(from_phone_purchase_date)

    End If

    If Not IsNull(to_phone_purchase_date) Then
        params.Add "to_phone_purchase_date", vba_datetime_to_api(to_phone_purchase_date)

    End If

    If Not IsNull(child_account_id) Then
        params.Add "child_account_id", serialize_list(child_account_id)

    End If

    If Not IsNull(children_phones_only) Then
        params.Add "children_phones_only", children_phones_only

    End If

    If Not IsNull(verification_name) Then
        params.Add "verification_name", verification_name

    End If

    If Not IsNull(verification_status) Then
        params.Add "verification_status", serialize_list(verification_status)

    End If

    If Not IsNull(from_unverified_hold_until) Then
        params.Add "from_unverified_hold_until", from_unverified_hold_until

    End If

    If Not IsNull(to_unverified_hold_until) Then
        params.Add "to_unverified_hold_until", to_unverified_hold_until

    End If

    If Not IsNull(can_be_used) Then
        params.Add "can_be_used", can_be_used

    End If

    If Not IsNull(order_by) Then
        params.Add "order_by", order_by

    End If

    If Not IsNull(sandbox) Then
        params.Add "sandbox", sandbox

    End If

    If Not IsNull(count) Then
        params.Add "count", count

    End If

    If Not IsNull(offset) Then
        params.Add "offset", offset

    End If

    If Not IsNull(phone_region_name) Then
        params.Add "phone_region_name", serialize_list(phone_region_name)

    End If

    If Not IsNull(rule_id) Then
        params.Add "rule_id", serialize_list(rule_id)

    End If

    If Not IsNull(rule_name) Then
        params.Add "rule_name", serialize_list(rule_name)

    End If

    If Not IsNull(is_bound_to_rule) Then
        params.Add "is_bound_to_rule", is_bound_to_rule

    End If

    Set res = makeRequest("GetPhoneNumbers", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetPhoneNumbers = res
End Function

Public Function GetNewPhoneNumbers(country_code, phone_category_name, phone_region_id, Optional country_state = Null, Optional count = Null, Optional offset = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "country_code", country_code

    params.Add "phone_category_name", phone_category_name

    params.Add "phone_region_id", phone_region_id



    If Not IsNull(country_state) Then
        params.Add "country_state", country_state

    End If

    If Not IsNull(count) Then
        params.Add "count", count

    End If

    If Not IsNull(offset) Then
        params.Add "offset", offset

    End If

    Set res = makeRequest("GetNewPhoneNumbers", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetNewPhoneNumbers = res
End Function

Public Function GetPhoneNumberCategories(Optional country_code = Null, Optional sandbox = Null, Optional locale = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(country_code) Then
        params.Add "country_code", country_code

    End If

    If Not IsNull(sandbox) Then
        params.Add "sandbox", sandbox

    End If

    If Not IsNull(locale) Then
        params.Add "locale", locale

    End If

    Set res = makeRequest("GetPhoneNumberCategories", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetPhoneNumberCategories = res
End Function

Public Function GetPhoneNumberCountryStates(country_code, phone_category_name, Optional country_state = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "country_code", country_code

    params.Add "phone_category_name", phone_category_name



    If Not IsNull(country_state) Then
        params.Add "country_state", country_state

    End If

    Set res = makeRequest("GetPhoneNumberCountryStates", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetPhoneNumberCountryStates = res
End Function

Public Function GetPhoneNumberRegions(country_code, phone_category_name, Optional country_state = Null, Optional omit_empty = Null, Optional phone_region_id = Null, Optional phone_region_name = Null, Optional phone_region_code = Null, Optional locale = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "country_code", country_code

    params.Add "phone_category_name", phone_category_name



    If Not IsNull(country_state) Then
        params.Add "country_state", country_state

    End If

    If Not IsNull(omit_empty) Then
        params.Add "omit_empty", omit_empty

    End If

    If Not IsNull(phone_region_id) Then
        params.Add "phone_region_id", phone_region_id

    End If

    If Not IsNull(phone_region_name) Then
        params.Add "phone_region_name", phone_region_name

    End If

    If Not IsNull(phone_region_code) Then
        params.Add "phone_region_code", phone_region_code

    End If

    If Not IsNull(locale) Then
        params.Add "locale", locale

    End If

    Set res = makeRequest("GetPhoneNumberRegions", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetPhoneNumberRegions = res
End Function

Public Function GetActualPhoneNumberRegion(country_code, phone_category_name, phone_region_id, Optional country_state = Null, Optional locale = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "country_code", country_code

    params.Add "phone_category_name", phone_category_name

    params.Add "phone_region_id", phone_region_id



    If Not IsNull(country_state) Then
        params.Add "country_state", country_state

    End If

    If Not IsNull(locale) Then
        params.Add "locale", locale

    End If

    Set res = makeRequest("GetActualPhoneNumberRegion", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetActualPhoneNumberRegion = res
End Function

Public Function AddCallerID(callerid_number) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "callerid_number", callerid_number



    Set res = makeRequest("AddCallerID", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set AddCallerID = res
End Function

Public Function ActivateCallerID(verification_code, Optional callerid_id = Null, Optional callerid_number = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "verification_code", verification_code



    If Not IsNull(callerid_id) Then
        params.Add "callerid_id", callerid_id

    End If

    If Not IsNull(callerid_number) Then
        params.Add "callerid_number", callerid_number

    End If

    Set res = makeRequest("ActivateCallerID", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set ActivateCallerID = res
End Function

Public Function DelCallerID(Optional callerid_id = Null, Optional callerid_number = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(callerid_id) Then
        params.Add "callerid_id", callerid_id

    End If

    If Not IsNull(callerid_number) Then
        params.Add "callerid_number", callerid_number

    End If

    Set res = makeRequest("DelCallerID", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set DelCallerID = res
End Function

Public Function GetCallerIDs(Optional callerid_id = Null, Optional callerid_number = Null, Optional active = Null, Optional order_by = Null, Optional count = Null, Optional offset = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(callerid_id) Then
        params.Add "callerid_id", callerid_id

    End If

    If Not IsNull(callerid_number) Then
        params.Add "callerid_number", callerid_number

    End If

    If Not IsNull(active) Then
        params.Add "active", active

    End If

    If Not IsNull(order_by) Then
        params.Add "order_by", order_by

    End If

    If Not IsNull(count) Then
        params.Add "count", count

    End If

    If Not IsNull(offset) Then
        params.Add "offset", offset

    End If

    Set res = makeRequest("GetCallerIDs", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetCallerIDs = res
End Function

Public Function VerifyCallerID(Optional callerid_id = Null, Optional callerid_number = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(callerid_id) Then
        params.Add "callerid_id", callerid_id

    End If

    If Not IsNull(callerid_number) Then
        params.Add "callerid_number", callerid_number

    End If

    Set res = makeRequest("VerifyCallerID", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set VerifyCallerID = res
End Function

Public Function AddQueue(acd_queue_name, Optional application_id = Null, Optional application_name = Null, Optional acd_queue_priority = Null, Optional auto_binding = Null, Optional service_probability = Null, Optional max_queue_size = Null, Optional max_waiting_time = Null, Optional average_service_time = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "acd_queue_name", acd_queue_name



    If Not IsNull(application_id) Then
        params.Add "application_id", application_id

    End If

    If Not IsNull(application_name) Then
        params.Add "application_name", application_name

    End If

    If Not IsNull(acd_queue_priority) Then
        params.Add "acd_queue_priority", acd_queue_priority

    End If

    If Not IsNull(auto_binding) Then
        params.Add "auto_binding", auto_binding

    End If

    If Not IsNull(service_probability) Then
        params.Add "service_probability", service_probability

    End If

    If Not IsNull(max_queue_size) Then
        params.Add "max_queue_size", max_queue_size

    End If

    If Not IsNull(max_waiting_time) Then
        params.Add "max_waiting_time", max_waiting_time

    End If

    If Not IsNull(average_service_time) Then
        params.Add "average_service_time", average_service_time

    End If

    Set res = makeRequest("AddQueue", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set AddQueue = res
End Function

Public Function BindUserToQueue(bind, Optional application_id = Null, Optional application_name = Null, Optional user_id = Null, Optional user_name = Null, Optional acd_queue_id = Null, Optional acd_queue_name = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "bind", bind



    If Not IsNull(application_id) Then
        params.Add "application_id", application_id

    End If

    If Not IsNull(application_name) Then
        params.Add "application_name", application_name

    End If

    If Not IsNull(user_id) Then
        params.Add "user_id", serialize_list(user_id)

    End If

    If Not IsNull(user_name) Then
        params.Add "user_name", serialize_list(user_name)

    End If

    If Not IsNull(acd_queue_id) Then
        params.Add "acd_queue_id", serialize_list(acd_queue_id)

    End If

    If Not IsNull(acd_queue_name) Then
        params.Add "acd_queue_name", serialize_list(acd_queue_name)

    End If

    Set res = makeRequest("BindUserToQueue", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set BindUserToQueue = res
End Function

Public Function DelQueue(Optional acd_queue_id = Null, Optional acd_queue_name = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(acd_queue_id) Then
        params.Add "acd_queue_id", serialize_list(acd_queue_id)

    End If

    If Not IsNull(acd_queue_name) Then
        params.Add "acd_queue_name", serialize_list(acd_queue_name)

    End If

    Set res = makeRequest("DelQueue", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set DelQueue = res
End Function

Public Function SetQueueInfo(Optional acd_queue_id = Null, Optional acd_queue_name = Null, Optional new_acd_queue_name = Null, Optional acd_queue_priority = Null, Optional auto_binding = Null, Optional service_probability = Null, Optional max_queue_size = Null, Optional max_waiting_time = Null, Optional average_service_time = Null, Optional application_id = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(acd_queue_id) Then
        params.Add "acd_queue_id", acd_queue_id

    End If

    If Not IsNull(acd_queue_name) Then
        params.Add "acd_queue_name", acd_queue_name

    End If

    If Not IsNull(new_acd_queue_name) Then
        params.Add "new_acd_queue_name", new_acd_queue_name

    End If

    If Not IsNull(acd_queue_priority) Then
        params.Add "acd_queue_priority", acd_queue_priority

    End If

    If Not IsNull(auto_binding) Then
        params.Add "auto_binding", auto_binding

    End If

    If Not IsNull(service_probability) Then
        params.Add "service_probability", service_probability

    End If

    If Not IsNull(max_queue_size) Then
        params.Add "max_queue_size", max_queue_size

    End If

    If Not IsNull(max_waiting_time) Then
        params.Add "max_waiting_time", max_waiting_time

    End If

    If Not IsNull(average_service_time) Then
        params.Add "average_service_time", average_service_time

    End If

    If Not IsNull(application_id) Then
        params.Add "application_id", application_id

    End If

    Set res = makeRequest("SetQueueInfo", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set SetQueueInfo = res
End Function

Public Function GetQueues(Optional acd_queue_id = Null, Optional acd_queue_name = Null, Optional application_id = Null, Optional skill_id = Null, Optional excluded_skill_id = Null, Optional with_skills = Null, Optional showing_skill_id = Null, Optional count = Null, Optional offset = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(acd_queue_id) Then
        params.Add "acd_queue_id", acd_queue_id

    End If

    If Not IsNull(acd_queue_name) Then
        params.Add "acd_queue_name", acd_queue_name

    End If

    If Not IsNull(application_id) Then
        params.Add "application_id", application_id

    End If

    If Not IsNull(skill_id) Then
        params.Add "skill_id", skill_id

    End If

    If Not IsNull(excluded_skill_id) Then
        params.Add "excluded_skill_id", excluded_skill_id

    End If

    If Not IsNull(with_skills) Then
        params.Add "with_skills", with_skills

    End If

    If Not IsNull(showing_skill_id) Then
        params.Add "showing_skill_id", showing_skill_id

    End If

    If Not IsNull(count) Then
        params.Add "count", count

    End If

    If Not IsNull(offset) Then
        params.Add "offset", offset

    End If

    Set res = makeRequest("GetQueues", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetQueues = res
End Function

Public Function GetACDState(Optional acd_queue_id = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(acd_queue_id) Then
        params.Add "acd_queue_id", serialize_list(acd_queue_id)

    End If

    Set res = makeRequest("GetACDState", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetACDState = res
End Function

Public Function GetACDOperatorStatistics(from_date, user_id, Optional to_date = Null, Optional acd_queue_id = Null, Optional abbreviation = Null, Optional report = Null, Optional aggregation = Null, Optional group = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "from_date", vba_datetime_to_api(from_date)

    params.Add "user_id", serialize_list(user_id)



    If Not IsNull(to_date) Then
        params.Add "to_date", vba_datetime_to_api(to_date)

    End If

    If Not IsNull(acd_queue_id) Then
        params.Add "acd_queue_id", serialize_list(acd_queue_id)

    End If

    If Not IsNull(abbreviation) Then
        params.Add "abbreviation", abbreviation

    End If

    If Not IsNull(report) Then
        params.Add "report", serialize_list(report)

    End If

    If Not IsNull(aggregation) Then
        params.Add "aggregation", aggregation

    End If

    If Not IsNull(group) Then
        params.Add "group", group

    End If

    Set res = makeRequest("GetACDOperatorStatistics", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetACDOperatorStatistics = res
End Function

Public Function GetACDQueueStatistics(from_date, Optional to_date = Null, Optional abbreviation = Null, Optional acd_queue_id = Null, Optional report = Null, Optional aggregation = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "from_date", vba_datetime_to_api(from_date)



    If Not IsNull(to_date) Then
        params.Add "to_date", vba_datetime_to_api(to_date)

    End If

    If Not IsNull(abbreviation) Then
        params.Add "abbreviation", abbreviation

    End If

    If Not IsNull(acd_queue_id) Then
        params.Add "acd_queue_id", serialize_list(acd_queue_id)

    End If

    If Not IsNull(report) Then
        params.Add "report", serialize_list(report)

    End If

    If Not IsNull(aggregation) Then
        params.Add "aggregation", aggregation

    End If

    Set res = makeRequest("GetACDQueueStatistics", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetACDQueueStatistics = res
End Function

Public Function GetACDOperatorStatusStatistics(from_date, user_id, Optional to_date = Null, Optional acd_status = Null, Optional aggregation = Null, Optional group = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "from_date", vba_datetime_to_api(from_date)

    params.Add "user_id", serialize_list(user_id)



    If Not IsNull(to_date) Then
        params.Add "to_date", vba_datetime_to_api(to_date)

    End If

    If Not IsNull(acd_status) Then
        params.Add "acd_status", serialize_list(acd_status)

    End If

    If Not IsNull(aggregation) Then
        params.Add "aggregation", aggregation

    End If

    If Not IsNull(group) Then
        params.Add "group", group

    End If

    Set res = makeRequest("GetACDOperatorStatusStatistics", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetACDOperatorStatusStatistics = res
End Function

Public Function AddSkill(skill_name) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "skill_name", skill_name



    Set res = makeRequest("AddSkill", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set AddSkill = res
End Function

Public Function DelSkill(Optional skill_id = Null, Optional skill_name = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(skill_id) Then
        params.Add "skill_id", skill_id

    End If

    If Not IsNull(skill_name) Then
        params.Add "skill_name", skill_name

    End If

    Set res = makeRequest("DelSkill", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set DelSkill = res
End Function

Public Function SetSkillInfo(new_skill_name, Optional skill_id = Null, Optional skill_name = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "new_skill_name", new_skill_name



    If Not IsNull(skill_id) Then
        params.Add "skill_id", skill_id

    End If

    If Not IsNull(skill_name) Then
        params.Add "skill_name", skill_name

    End If

    Set res = makeRequest("SetSkillInfo", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set SetSkillInfo = res
End Function

Public Function GetSkills(Optional skill_id = Null, Optional skill_name = Null, Optional count = Null, Optional offset = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(skill_id) Then
        params.Add "skill_id", skill_id

    End If

    If Not IsNull(skill_name) Then
        params.Add "skill_name", skill_name

    End If

    If Not IsNull(count) Then
        params.Add "count", count

    End If

    If Not IsNull(offset) Then
        params.Add "offset", offset

    End If

    Set res = makeRequest("GetSkills", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetSkills = res
End Function

Public Function BindSkill(Optional skill_id = Null, Optional skill_name = Null, Optional user_id = Null, Optional user_name = Null, Optional acd_queue_id = Null, Optional acd_queue_name = Null, Optional application_id = Null, Optional application_name = Null, Optional bind = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(skill_id) Then
        params.Add "skill_id", serialize_list(skill_id)

    End If

    If Not IsNull(skill_name) Then
        params.Add "skill_name", serialize_list(skill_name)

    End If

    If Not IsNull(user_id) Then
        params.Add "user_id", serialize_list(user_id)

    End If

    If Not IsNull(user_name) Then
        params.Add "user_name", serialize_list(user_name)

    End If

    If Not IsNull(acd_queue_id) Then
        params.Add "acd_queue_id", serialize_list(acd_queue_id)

    End If

    If Not IsNull(acd_queue_name) Then
        params.Add "acd_queue_name", serialize_list(acd_queue_name)

    End If

    If Not IsNull(application_id) Then
        params.Add "application_id", application_id

    End If

    If Not IsNull(application_name) Then
        params.Add "application_name", application_name

    End If

    If Not IsNull(bind) Then
        params.Add "bind", bind

    End If

    Set res = makeRequest("BindSkill", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set BindSkill = res
End Function

Public Function GetAccountDocuments(Optional with_details = Null, Optional verification_name = Null, Optional verification_status = Null, Optional from_unverified_hold_until = Null, Optional to_unverified_hold_until = Null, Optional child_account_id = Null, Optional children_verifications_only = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(with_details) Then
        params.Add "with_details", with_details

    End If

    If Not IsNull(verification_name) Then
        params.Add "verification_name", verification_name

    End If

    If Not IsNull(verification_status) Then
        params.Add "verification_status", serialize_list(verification_status)

    End If

    If Not IsNull(from_unverified_hold_until) Then
        params.Add "from_unverified_hold_until", from_unverified_hold_until

    End If

    If Not IsNull(to_unverified_hold_until) Then
        params.Add "to_unverified_hold_until", to_unverified_hold_until

    End If

    If Not IsNull(child_account_id) Then
        params.Add "child_account_id", serialize_list(child_account_id)

    End If

    If Not IsNull(children_verifications_only) Then
        params.Add "children_verifications_only", children_verifications_only

    End If

    Set res = makeRequest("GetAccountDocuments", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetAccountDocuments = res
End Function

Public Function AddAdminUser(new_admin_user_name, admin_user_display_name, new_admin_user_password, Optional admin_user_active = Null, Optional admin_role_id = Null, Optional admin_role_name = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "new_admin_user_name", new_admin_user_name

    params.Add "admin_user_display_name", admin_user_display_name

    params.Add "new_admin_user_password", new_admin_user_password



    If Not IsNull(admin_user_active) Then
        params.Add "admin_user_active", admin_user_active

    End If

    If Not IsNull(admin_role_id) Then
        params.Add "admin_role_id", admin_role_id

    End If

    If Not IsNull(admin_role_name) Then
        params.Add "admin_role_name", serialize_list(admin_role_name)

    End If

    Set res = makeRequest("AddAdminUser", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set AddAdminUser = res
End Function

Public Function DelAdminUser(Optional required_admin_user_id = Null, Optional required_admin_user_name = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(required_admin_user_id) Then
        params.Add "required_admin_user_id", serialize_list(required_admin_user_id)

    End If

    If Not IsNull(required_admin_user_name) Then
        params.Add "required_admin_user_name", serialize_list(required_admin_user_name)

    End If

    Set res = makeRequest("DelAdminUser", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set DelAdminUser = res
End Function

Public Function SetAdminUserInfo(Optional required_admin_user_id = Null, Optional required_admin_user_name = Null, Optional new_admin_user_name = Null, Optional admin_user_display_name = Null, Optional new_admin_user_password = Null, Optional admin_user_active = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(required_admin_user_id) Then
        params.Add "required_admin_user_id", required_admin_user_id

    End If

    If Not IsNull(required_admin_user_name) Then
        params.Add "required_admin_user_name", required_admin_user_name

    End If

    If Not IsNull(new_admin_user_name) Then
        params.Add "new_admin_user_name", new_admin_user_name

    End If

    If Not IsNull(admin_user_display_name) Then
        params.Add "admin_user_display_name", admin_user_display_name

    End If

    If Not IsNull(new_admin_user_password) Then
        params.Add "new_admin_user_password", new_admin_user_password

    End If

    If Not IsNull(admin_user_active) Then
        params.Add "admin_user_active", admin_user_active

    End If

    Set res = makeRequest("SetAdminUserInfo", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set SetAdminUserInfo = res
End Function

Public Function GetAdminUsers(Optional required_admin_user_id = Null, Optional required_admin_user_name = Null, Optional admin_user_display_name = Null, Optional admin_user_active = Null, Optional with_roles = Null, Optional with_access_entries = Null, Optional count = Null, Optional offset = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(required_admin_user_id) Then
        params.Add "required_admin_user_id", required_admin_user_id

    End If

    If Not IsNull(required_admin_user_name) Then
        params.Add "required_admin_user_name", required_admin_user_name

    End If

    If Not IsNull(admin_user_display_name) Then
        params.Add "admin_user_display_name", admin_user_display_name

    End If

    If Not IsNull(admin_user_active) Then
        params.Add "admin_user_active", admin_user_active

    End If

    If Not IsNull(with_roles) Then
        params.Add "with_roles", with_roles

    End If

    If Not IsNull(with_access_entries) Then
        params.Add "with_access_entries", with_access_entries

    End If

    If Not IsNull(count) Then
        params.Add "count", count

    End If

    If Not IsNull(offset) Then
        params.Add "offset", offset

    End If

    Set res = makeRequest("GetAdminUsers", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetAdminUsers = res
End Function

Public Function AddAdminRole(admin_role_name, Optional admin_role_active = Null, Optional like_admin_role_id = Null, Optional like_admin_role_name = Null, Optional allowed_entries = Null, Optional denied_entries = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "admin_role_name", admin_role_name



    If Not IsNull(admin_role_active) Then
        params.Add "admin_role_active", admin_role_active

    End If

    If Not IsNull(like_admin_role_id) Then
        params.Add "like_admin_role_id", serialize_list(like_admin_role_id)

    End If

    If Not IsNull(like_admin_role_name) Then
        params.Add "like_admin_role_name", serialize_list(like_admin_role_name)

    End If

    If Not IsNull(allowed_entries) Then
        params.Add "allowed_entries", serialize_list(allowed_entries)

    End If

    If Not IsNull(denied_entries) Then
        params.Add "denied_entries", serialize_list(denied_entries)

    End If

    Set res = makeRequest("AddAdminRole", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set AddAdminRole = res
End Function

Public Function DelAdminRole(Optional admin_role_id = Null, Optional admin_role_name = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(admin_role_id) Then
        params.Add "admin_role_id", serialize_list(admin_role_id)

    End If

    If Not IsNull(admin_role_name) Then
        params.Add "admin_role_name", serialize_list(admin_role_name)

    End If

    Set res = makeRequest("DelAdminRole", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set DelAdminRole = res
End Function

Public Function SetAdminRoleInfo(Optional admin_role_id = Null, Optional admin_role_name = Null, Optional new_admin_role_name = Null, Optional admin_role_active = Null, Optional entry_modification_mode = Null, Optional allowed_entries = Null, Optional denied_entries = Null, Optional like_admin_role_id = Null, Optional like_admin_role_name = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(admin_role_id) Then
        params.Add "admin_role_id", admin_role_id

    End If

    If Not IsNull(admin_role_name) Then
        params.Add "admin_role_name", admin_role_name

    End If

    If Not IsNull(new_admin_role_name) Then
        params.Add "new_admin_role_name", new_admin_role_name

    End If

    If Not IsNull(admin_role_active) Then
        params.Add "admin_role_active", admin_role_active

    End If

    If Not IsNull(entry_modification_mode) Then
        params.Add "entry_modification_mode", entry_modification_mode

    End If

    If Not IsNull(allowed_entries) Then
        params.Add "allowed_entries", serialize_list(allowed_entries)

    End If

    If Not IsNull(denied_entries) Then
        params.Add "denied_entries", serialize_list(denied_entries)

    End If

    If Not IsNull(like_admin_role_id) Then
        params.Add "like_admin_role_id", serialize_list(like_admin_role_id)

    End If

    If Not IsNull(like_admin_role_name) Then
        params.Add "like_admin_role_name", serialize_list(like_admin_role_name)

    End If

    Set res = makeRequest("SetAdminRoleInfo", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set SetAdminRoleInfo = res
End Function

Public Function GetAdminRoles(Optional admin_role_id = Null, Optional admin_role_name = Null, Optional admin_role_active = Null, Optional with_entries = Null, Optional with_account_roles = Null, Optional with_parent_roles = Null, Optional included_admin_user_id = Null, Optional excluded_admin_user_id = Null, Optional full_admin_users_matching = Null, Optional showing_admin_user_id = Null, Optional count = Null, Optional offset = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(admin_role_id) Then
        params.Add "admin_role_id", admin_role_id

    End If

    If Not IsNull(admin_role_name) Then
        params.Add "admin_role_name", admin_role_name

    End If

    If Not IsNull(admin_role_active) Then
        params.Add "admin_role_active", admin_role_active

    End If

    If Not IsNull(with_entries) Then
        params.Add "with_entries", with_entries

    End If

    If Not IsNull(with_account_roles) Then
        params.Add "with_account_roles", with_account_roles

    End If

    If Not IsNull(with_parent_roles) Then
        params.Add "with_parent_roles", with_parent_roles

    End If

    If Not IsNull(included_admin_user_id) Then
        params.Add "included_admin_user_id", serialize_list(included_admin_user_id)

    End If

    If Not IsNull(excluded_admin_user_id) Then
        params.Add "excluded_admin_user_id", serialize_list(excluded_admin_user_id)

    End If

    If Not IsNull(full_admin_users_matching) Then
        params.Add "full_admin_users_matching", full_admin_users_matching

    End If

    If Not IsNull(showing_admin_user_id) Then
        params.Add "showing_admin_user_id", showing_admin_user_id

    End If

    If Not IsNull(count) Then
        params.Add "count", count

    End If

    If Not IsNull(offset) Then
        params.Add "offset", offset

    End If

    Set res = makeRequest("GetAdminRoles", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetAdminRoles = res
End Function

Public Function AttachAdminRole(Optional required_admin_user_id = Null, Optional required_admin_user_name = Null, Optional admin_role_id = Null, Optional admin_role_name = Null, Optional mode = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(required_admin_user_id) Then
        params.Add "required_admin_user_id", serialize_list(required_admin_user_id)

    End If

    If Not IsNull(required_admin_user_name) Then
        params.Add "required_admin_user_name", serialize_list(required_admin_user_name)

    End If

    If Not IsNull(admin_role_id) Then
        params.Add "admin_role_id", serialize_list(admin_role_id)

    End If

    If Not IsNull(admin_role_name) Then
        params.Add "admin_role_name", serialize_list(admin_role_name)

    End If

    If Not IsNull(mode) Then
        params.Add "mode", mode

    End If

    Set res = makeRequest("AttachAdminRole", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set AttachAdminRole = res
End Function

Public Function GetAvailableAdminRoleEntries() As Object

    Dim params As New Dictionary
    Dim res As Object


    Set res = makeRequest("GetAvailableAdminRoleEntries", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetAvailableAdminRoleEntries = res
End Function

Public Function AddAuthorizedAccountIP(authorized_ip, Optional allowed = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "authorized_ip", authorized_ip



    If Not IsNull(allowed) Then
        params.Add "allowed", allowed

    End If

    Set res = makeRequest("AddAuthorizedAccountIP", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set AddAuthorizedAccountIP = res
End Function

Public Function DelAuthorizedAccountIP(Optional authorized_ip = Null, Optional contains_ip = Null, Optional allowed = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(authorized_ip) Then
        params.Add "authorized_ip", authorized_ip

    End If

    If Not IsNull(contains_ip) Then
        params.Add "contains_ip", contains_ip

    End If

    If Not IsNull(allowed) Then
        params.Add "allowed", allowed

    End If

    Set res = makeRequest("DelAuthorizedAccountIP", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set DelAuthorizedAccountIP = res
End Function

Public Function GetAuthorizedAccountIPs(Optional authorized_ip = Null, Optional allowed = Null, Optional contains_ip = Null, Optional count = Null, Optional offset = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(authorized_ip) Then
        params.Add "authorized_ip", authorized_ip

    End If

    If Not IsNull(allowed) Then
        params.Add "allowed", allowed

    End If

    If Not IsNull(contains_ip) Then
        params.Add "contains_ip", contains_ip

    End If

    If Not IsNull(count) Then
        params.Add "count", count

    End If

    If Not IsNull(offset) Then
        params.Add "offset", offset

    End If

    Set res = makeRequest("GetAuthorizedAccountIPs", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetAuthorizedAccountIPs = res
End Function

Public Function CheckAuthorizedAccountIP(authorized_ip) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "authorized_ip", authorized_ip



    Set res = makeRequest("CheckAuthorizedAccountIP", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set CheckAuthorizedAccountIP = res
End Function

Public Function LinkRegulationAddress(regulation_address_id, Optional phone_id = Null, Optional phone_number = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "regulation_address_id", regulation_address_id



    If Not IsNull(phone_id) Then
        params.Add "phone_id", phone_id

    End If

    If Not IsNull(phone_number) Then
        params.Add "phone_number", phone_number

    End If

    Set res = makeRequest("LinkRegulationAddress", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set LinkRegulationAddress = res
End Function

Public Function GetZIPCodes(country_code, Optional phone_region_code = Null, Optional count = Null, Optional offset = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "country_code", country_code



    If Not IsNull(phone_region_code) Then
        params.Add "phone_region_code", phone_region_code

    End If

    If Not IsNull(count) Then
        params.Add "count", count

    End If

    If Not IsNull(offset) Then
        params.Add "offset", offset

    End If

    Set res = makeRequest("GetZIPCodes", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetZIPCodes = res
End Function

Public Function GetRegulationsAddress(Optional country_code = Null, Optional phone_category_name = Null, Optional phone_region_code = Null, Optional regulation_address_id = Null, Optional verified = Null, Optional in_progress = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(country_code) Then
        params.Add "country_code", country_code

    End If

    If Not IsNull(phone_category_name) Then
        params.Add "phone_category_name", phone_category_name

    End If

    If Not IsNull(phone_region_code) Then
        params.Add "phone_region_code", phone_region_code

    End If

    If Not IsNull(regulation_address_id) Then
        params.Add "regulation_address_id", regulation_address_id

    End If

    If Not IsNull(verified) Then
        params.Add "verified", verified

    End If

    If Not IsNull(in_progress) Then
        params.Add "in_progress", in_progress

    End If

    Set res = makeRequest("GetRegulationsAddress", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetRegulationsAddress = res
End Function

Public Function GetAvailableRegulations(country_code, phone_category_name, Optional phone_region_code = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "country_code", country_code

    params.Add "phone_category_name", phone_category_name



    If Not IsNull(phone_region_code) Then
        params.Add "phone_region_code", phone_region_code

    End If

    Set res = makeRequest("GetAvailableRegulations", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetAvailableRegulations = res
End Function

Public Function GetCountries(Optional country_code = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(country_code) Then
        params.Add "country_code", country_code

    End If

    Set res = makeRequest("GetCountries", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetCountries = res
End Function

Public Function GetRegions(country_code, phone_category_name, Optional city_name = Null, Optional count = Null, Optional offset = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "country_code", country_code

    params.Add "phone_category_name", phone_category_name



    If Not IsNull(city_name) Then
        params.Add "city_name", city_name

    End If

    If Not IsNull(count) Then
        params.Add "count", count

    End If

    If Not IsNull(offset) Then
        params.Add "offset", offset

    End If

    Set res = makeRequest("GetRegions", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetRegions = res
End Function

Public Function AddPushCredential(Optional push_provider_name = Null, Optional push_provider_id = Null, Optional application_id = Null, Optional application_name = Null, Optional credential_bundle = Null, Optional cert_content = Null, Optional cert_password = Null, Optional is_dev_mode = Null, Optional sender_id = Null, Optional server_key = Null, Optional service_account_file = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(push_provider_name) Then
        params.Add "push_provider_name", push_provider_name

    End If

    If Not IsNull(push_provider_id) Then
        params.Add "push_provider_id", push_provider_id

    End If

    If Not IsNull(application_id) Then
        params.Add "application_id", application_id

    End If

    If Not IsNull(application_name) Then
        params.Add "application_name", application_name

    End If

    If Not IsNull(credential_bundle) Then
        params.Add "credential_bundle", credential_bundle

    End If

    If Not IsNull(cert_content) Then
        params.Add "cert_content", cert_content

    End If

    If Not IsNull(cert_password) Then
        params.Add "cert_password", cert_password

    End If

    If Not IsNull(is_dev_mode) Then
        params.Add "is_dev_mode", is_dev_mode

    End If

    If Not IsNull(sender_id) Then
        params.Add "sender_id", sender_id

    End If

    If Not IsNull(server_key) Then
        params.Add "server_key", server_key

    End If

    If Not IsNull(service_account_file) Then
        params.Add "service_account_file", service_account_file

    End If

    Set res = makeRequest("AddPushCredential", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set AddPushCredential = res
End Function

Public Function SetPushCredential(push_credential_id, Optional cert_content = Null, Optional cert_password = Null, Optional is_dev_mode = Null, Optional sender_id = Null, Optional server_key = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "push_credential_id", push_credential_id



    If Not IsNull(cert_content) Then
        params.Add "cert_content", cert_content

    End If

    If Not IsNull(cert_password) Then
        params.Add "cert_password", cert_password

    End If

    If Not IsNull(is_dev_mode) Then
        params.Add "is_dev_mode", is_dev_mode

    End If

    If Not IsNull(sender_id) Then
        params.Add "sender_id", sender_id

    End If

    If Not IsNull(server_key) Then
        params.Add "server_key", server_key

    End If

    Set res = makeRequest("SetPushCredential", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set SetPushCredential = res
End Function

Public Function DelPushCredential(push_credential_id) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "push_credential_id", push_credential_id



    Set res = makeRequest("DelPushCredential", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set DelPushCredential = res
End Function

Public Function GetPushCredential(Optional push_credential_id = Null, Optional push_provider_name = Null, Optional push_provider_id = Null, Optional application_name = Null, Optional application_id = Null, Optional with_cert = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(push_credential_id) Then
        params.Add "push_credential_id", push_credential_id

    End If

    If Not IsNull(push_provider_name) Then
        params.Add "push_provider_name", push_provider_name

    End If

    If Not IsNull(push_provider_id) Then
        params.Add "push_provider_id", push_provider_id

    End If

    If Not IsNull(application_name) Then
        params.Add "application_name", application_name

    End If

    If Not IsNull(application_id) Then
        params.Add "application_id", application_id

    End If

    If Not IsNull(with_cert) Then
        params.Add "with_cert", with_cert

    End If

    Set res = makeRequest("GetPushCredential", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetPushCredential = res
End Function

Public Function BindPushCredential(push_credential_id, application_id, Optional bind = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "push_credential_id", serialize_list(push_credential_id)

    params.Add "application_id", serialize_list(application_id)



    If Not IsNull(bind) Then
        params.Add "bind", bind

    End If

    Set res = makeRequest("BindPushCredential", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set BindPushCredential = res
End Function

Public Function AddDialogflowKey(application_id, json_credentials, Optional application_name = Null, Optional description = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "application_id", application_id

    params.Add "json_credentials", json_credentials



    If Not IsNull(application_name) Then
        params.Add "application_name", application_name

    End If

    If Not IsNull(description) Then
        params.Add "description", description

    End If

    Set res = makeRequest("AddDialogflowKey", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set AddDialogflowKey = res
End Function

Public Function SetDialogflowKey(dialogflow_key_id, description) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "dialogflow_key_id", dialogflow_key_id

    params.Add "description", description



    Set res = makeRequest("SetDialogflowKey", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set SetDialogflowKey = res
End Function

Public Function DelDialogflowKey(dialogflow_key_id) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "dialogflow_key_id", dialogflow_key_id



    Set res = makeRequest("DelDialogflowKey", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set DelDialogflowKey = res
End Function

Public Function GetDialogflowKeys(Optional dialogflow_key_id = Null, Optional application_name = Null, Optional application_id = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(dialogflow_key_id) Then
        params.Add "dialogflow_key_id", dialogflow_key_id

    End If

    If Not IsNull(application_name) Then
        params.Add "application_name", application_name

    End If

    If Not IsNull(application_id) Then
        params.Add "application_id", application_id

    End If

    Set res = makeRequest("GetDialogflowKeys", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetDialogflowKeys = res
End Function

Public Function BindDialogflowKeys(dialogflow_key_id, application_id, Optional bind = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "dialogflow_key_id", dialogflow_key_id

    params.Add "application_id", serialize_list(application_id)



    If Not IsNull(bind) Then
        params.Add "bind", bind

    End If

    Set res = makeRequest("BindDialogflowKeys", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set BindDialogflowKeys = res
End Function

Public Function SendSmsMessage(source, destination, sms_body) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "source", source

    params.Add "destination", destination

    params.Add "sms_body", sms_body



    Set res = makeRequest("SendSmsMessage", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set SendSmsMessage = res
End Function

Public Function A2PSendSms(src_number, dst_numbers, text) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "src_number", src_number

    params.Add "dst_numbers", serialize_list(dst_numbers)

    params.Add "text", text



    Set res = makeRequest("A2PSendSms", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set A2PSendSms = res
End Function

Public Function ControlSms(phone_number, command) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "phone_number", phone_number

    params.Add "command", command



    Set res = makeRequest("ControlSms", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set ControlSms = res
End Function

Public Function GetRecordStorages(Optional record_storage_id = Null, Optional record_storage_name = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(record_storage_id) Then
        params.Add "record_storage_id", serialize_list(record_storage_id)

    End If

    If Not IsNull(record_storage_name) Then
        params.Add "record_storage_name", serialize_list(record_storage_name)

    End If

    Set res = makeRequest("GetRecordStorages", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetRecordStorages = res
End Function

Public Function CreateKey(Optional description = Null, Optional role_id = Null, Optional role_name = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(description) Then
        params.Add "description", description

    End If

    If Not IsNull(role_id) Then
        params.Add "role_id", serialize_list(role_id)

    End If

    If Not IsNull(role_name) Then
        params.Add "role_name", serialize_list(role_name)

    End If

    Set res = makeRequest("CreateKey", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set CreateKey = res
End Function

Public Function GetKeys(Optional key_id = Null, Optional with_roles = Null, Optional offset = Null, Optional count = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(key_id) Then
        params.Add "key_id", key_id

    End If

    If Not IsNull(with_roles) Then
        params.Add "with_roles", with_roles

    End If

    If Not IsNull(offset) Then
        params.Add "offset", offset

    End If

    If Not IsNull(count) Then
        params.Add "count", count

    End If

    Set res = makeRequest("GetKeys", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetKeys = res
End Function

Public Function UpdateKey(key_id, description) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "key_id", key_id

    params.Add "description", description



    Set res = makeRequest("UpdateKey", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set UpdateKey = res
End Function

Public Function DeleteKey(key_id) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "key_id", key_id



    Set res = makeRequest("DeleteKey", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set DeleteKey = res
End Function

Public Function SetKeyRoles(key_id, Optional role_id = Null, Optional role_name = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "key_id", key_id



    If Not IsNull(role_id) Then
        params.Add "role_id", serialize_list(role_id)

    End If

    If Not IsNull(role_name) Then
        params.Add "role_name", serialize_list(role_name)

    End If

    Set res = makeRequest("SetKeyRoles", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set SetKeyRoles = res
End Function

Public Function GetKeyRoles(key_id, Optional with_expanded_roles = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "key_id", key_id



    If Not IsNull(with_expanded_roles) Then
        params.Add "with_expanded_roles", with_expanded_roles

    End If

    Set res = makeRequest("GetKeyRoles", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetKeyRoles = res
End Function

Public Function RemoveKeyRoles(key_id, Optional role_id = Null, Optional role_name = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "key_id", key_id



    If Not IsNull(role_id) Then
        params.Add "role_id", serialize_list(role_id)

    End If

    If Not IsNull(role_name) Then
        params.Add "role_name", serialize_list(role_name)

    End If

    Set res = makeRequest("RemoveKeyRoles", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set RemoveKeyRoles = res
End Function

Public Function AddSubUser(new_subuser_name, new_subuser_password, Optional role_id = Null, Optional role_name = Null, Optional description = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "new_subuser_name", new_subuser_name

    params.Add "new_subuser_password", new_subuser_password



    If Not IsNull(role_id) Then
        params.Add "role_id", serialize_list(role_id)

    End If

    If Not IsNull(role_name) Then
        params.Add "role_name", serialize_list(role_name)

    End If

    If Not IsNull(description) Then
        params.Add "description", description

    End If

    Set res = makeRequest("AddSubUser", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set AddSubUser = res
End Function

Public Function GetSubUsers(Optional subuser_id = Null, Optional with_roles = Null, Optional offset = Null, Optional count = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(subuser_id) Then
        params.Add "subuser_id", subuser_id

    End If

    If Not IsNull(with_roles) Then
        params.Add "with_roles", with_roles

    End If

    If Not IsNull(offset) Then
        params.Add "offset", offset

    End If

    If Not IsNull(count) Then
        params.Add "count", count

    End If

    Set res = makeRequest("GetSubUsers", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetSubUsers = res
End Function

Public Function SetSubUserInfo(subuser_id, Optional old_subuser_password = Null, Optional new_subuser_password = Null, Optional description = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "subuser_id", subuser_id



    If Not IsNull(old_subuser_password) Then
        params.Add "old_subuser_password", old_subuser_password

    End If

    If Not IsNull(new_subuser_password) Then
        params.Add "new_subuser_password", new_subuser_password

    End If

    If Not IsNull(description) Then
        params.Add "description", description

    End If

    Set res = makeRequest("SetSubUserInfo", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set SetSubUserInfo = res
End Function

Public Function DelSubUser(subuser_id) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "subuser_id", subuser_id



    Set res = makeRequest("DelSubUser", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set DelSubUser = res
End Function

Public Function SetSubUserRoles(subuser_id, Optional role_id = Null, Optional role_name = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "subuser_id", subuser_id



    If Not IsNull(role_id) Then
        params.Add "role_id", serialize_list(role_id)

    End If

    If Not IsNull(role_name) Then
        params.Add "role_name", serialize_list(role_name)

    End If

    Set res = makeRequest("SetSubUserRoles", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set SetSubUserRoles = res
End Function

Public Function GetSubUserRoles(subuser_id, Optional with_expanded_roles = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "subuser_id", subuser_id



    If Not IsNull(with_expanded_roles) Then
        params.Add "with_expanded_roles", with_expanded_roles

    End If

    Set res = makeRequest("GetSubUserRoles", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetSubUserRoles = res
End Function

Public Function RemoveSubUserRoles(subuser_id, Optional role_id = Null, Optional role_name = Null, Optional force = Null) As Object

    Dim params As New Dictionary
    Dim res As Object
    params.Add "subuser_id", subuser_id



    If Not IsNull(role_id) Then
        params.Add "role_id", serialize_list(role_id)

    End If

    If Not IsNull(role_name) Then
        params.Add "role_name", serialize_list(role_name)

    End If

    If Not IsNull(force) Then
        params.Add "force", force

    End If

    Set res = makeRequest("RemoveSubUserRoles", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set RemoveSubUserRoles = res
End Function

Public Function GetRoles(Optional group_name = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(group_name) Then
        params.Add "group_name", group_name

    End If

    Set res = makeRequest("GetRoles", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetRoles = res
End Function

Public Function GetRoleGroups() As Object

    Dim params As New Dictionary
    Dim res As Object


    Set res = makeRequest("GetRoleGroups", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetRoleGroups = res
End Function

Public Function GetSmsHistory(Optional source_number = Null, Optional destination_number = Null, Optional direction = Null, Optional count = Null, Optional offset = Null, Optional from_date = Null, Optional to_date = Null, Optional output = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(source_number) Then
        params.Add "source_number", source_number

    End If

    If Not IsNull(destination_number) Then
        params.Add "destination_number", destination_number

    End If

    If Not IsNull(direction) Then
        params.Add "direction", direction

    End If

    If Not IsNull(count) Then
        params.Add "count", count

    End If

    If Not IsNull(offset) Then
        params.Add "offset", offset

    End If

    If Not IsNull(from_date) Then
        params.Add "from_date", vba_datetime_to_api(from_date)

    End If

    If Not IsNull(to_date) Then
        params.Add "to_date", vba_datetime_to_api(to_date)

    End If

    If Not IsNull(output) Then
        params.Add "output", output

    End If

    Set res = makeRequest("GetSmsHistory", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set GetSmsHistory = res
End Function

Public Function A2PGetSmsHistory(Optional source_number = Null, Optional destination_number = Null, Optional count = Null, Optional offset = Null, Optional from_date = Null, Optional to_date = Null, Optional output = Null, Optional delivery_status = Null) As Object

    Dim params As New Dictionary
    Dim res As Object


    If Not IsNull(source_number) Then
        params.Add "source_number", source_number

    End If

    If Not IsNull(destination_number) Then
        params.Add "destination_number", destination_number

    End If

    If Not IsNull(count) Then
        params.Add "count", count

    End If

    If Not IsNull(offset) Then
        params.Add "offset", offset

    End If

    If Not IsNull(from_date) Then
        params.Add "from_date", vba_datetime_to_api(from_date)

    End If

    If Not IsNull(to_date) Then
        params.Add "to_date", vba_datetime_to_api(to_date)

    End If

    If Not IsNull(output) Then
        params.Add "output", output

    End If

    If Not IsNull(delivery_status) Then
        params.Add "delivery_status", delivery_status

    End If

    Set res = makeRequest("A2PGetSmsHistory", params)
    If res.Exists("error") Then
        Debug.Print("ERROR: " & res("error")("msg"))
    End If
    Set A2PGetSmsHistory = res
End Function
