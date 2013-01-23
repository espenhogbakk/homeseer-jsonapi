<%@ LANGUAGE=VBScript %>
<%
' *****************************************************************************
dim i, j, k, l, dev, s, qaction, qid, qvalue, id
dim locations(64)
dim numlocations
dim opt_showhidden
dim jsonapi_version

jsonapi_version = "0.1"


' Get action from querystring
qaction = request.querystring("action")
if qaction = "" then qaction = "none"

' Get identifier
qid = request.querystring("id")
if qid = "" then qid = "none"

' Get value
qvalue = request.querystring("value")
if qvalue = "" then qvalue = "none"


' Get show hidden flag
if request.querystring("showhidden") <> "" then opt_showhidden = lcase(request.querystring("showhidden"))
if opt_showhidden <> "yes" then opt_showhidden = "no"

' build a list of locations
BuildLocationsList

' handle actions, checking for authorizations if needed
if qaction <> "none" then
    if GetUserAuthorizations(hs.WebLoggedInUser) <= 1 then
  	    response.write "Sorry, control is available to authorized users only!"
        hs.WriteLog "JSONAPI","Ignoring action from " & hs.WebLoggedInUser & " at " & request.ServerVariables("REMOTE_ADDR")
		  else
        HandleAction qaction, qid, qvalue
      end if
  end if


' *****************************************************************************
sub HandleAction(qaction, qid, qvalue)
  ' execute an action; we're assuming that guests have already been filtered out
  s = lcase(qaction)

  if s = "getrooms" then
    ShowRooms

  elseif s = "getroom" then
    ShowRoom qid

  elseif s = "getdevices" then
    ShowDevices

  elseif s = "getdevice" then
    ShowDevice qid

  elseif s = "deviceon" then
    DeviceOn qid

  elseif s = "deviceoff" then
    DeviceOff qid

  elseif s = "setdevicevalue" then
    SetDeviceValue qid, qvalue

  elseif s = "getevents" then
    ShowEvents

  elseif s = "runevent" then
    RunEvent qid
  
  else ' unknown action
    response.write "Unknown action " & qaction & " ignored."
    hs.WriteLog "JSONAPI","Ignoring unknown action " & qaction & " from " & hs.WebLoggedInUser & " at " & request.ServerVariables("REMOTE_ADDR")
  end if
end sub

' end of main routine -- everything else is functions and subs


' *****************************************************************************
sub ShowRooms()
  dim count, i, s

  ' show a list of locations as links to those locations
  response.write "["
  
  for i = 1 to numlocations
    response.write "{"
    response.write """id"": """ & lcase(EncodeString(locations(i))) & """, "
    response.write """name"": """ & titleCase(locations(i)) & """"
    response.write "}"
    if not i = numlocations then
      response.write ", "
    end if
    next

  count = 0

  response.write "]"
end sub


' *****************************************************************************
sub ShowRoom(room)
  dim dev, devices, devname

  response.write "{""id"": """ & lcase(room) & """, "
  response.write """name"": """ & titleCase(Replace(DecodeString(room), "_", " "))  & """, "
  response.write """devices"": ["

  j = hs.DeviceCount

  thisroom = lcase(Replace(room,"_"," "))
  if thisroom = "no_location" then thisroom = ""

  set devices = hs.GetDeviceEnumerator

  ' loop through all devices
  if IsObject(devices) then
      do while not devices.Finished
        set dev = devices.GetNext
        if not dev is nothing then
            if (opt_showhidden = "yes" or (dev.misc and &H20) = 0) and (lcase(dev.location) = thisroom or thisroom = "all") then
                ' show this device
                id = dev.hc & dev.dc
                ShowDevice(id)
                response.write ", "
              end if
          end if
        loop
    end if

  response.write "]}"
end sub


' *****************************************************************************
sub ShowDevices()
  dim devices
  set devices = hs.GetDeviceEnumerator

  response.write "["

  ' loop through all devices
  if IsObject(devices) then
      do while not devices.Finished
        set dev = devices.GetNext
        if not dev is nothing then
            if (opt_showhidden = "yes" or (dev.misc and &H20) = 0) then
                ' show this device
                id = dev.hc & dev.dc
                ShowDevice(id)
                response.write ", "
              end if
          end if
        loop
    end if

  response.write "]"

end sub


' *****************************************************************************
sub ShowDevice(id)
  'Render a device
  dim device
  set device = hs.GetDeviceByRef(hs.DeviceExistsRef(id))

  deviceStatus = hs.DeviceString(id)
  if deviceStatus = "" then
    select case hs.DeviceStatus(id)
      case 2
        deviceStatus = "true"
      case 3
        deviceStatus = "false"
      case 4
        deviceStatus = "true"
      case else
        deviceStatus = "unknown"
      end select
  end if

  deviceValue = hs.DeviceValue(id)

  response.write "{"
  response.write """id"": """ & id & """, "
  response.write """name"": """ & device.name & """, "
  response.write """room"": """ & device.location & """, "
  response.write """floor"": """ & device.location2 & """, "
  response.write """dimmable"": " & lcase(device.can_dim) & ", "
  response.write """on"": " & lcase(deviceStatus) & ", "
  response.write """value"": " & deviceValue & ", "
  response.write """since"": """ & hs.DeviceLastChange(id) & """, "
  response.write """type"": """ & device.dev_type_string & """, "
  response.write """status_support"": " & lcase(device.status_support) & ", "
  response.write """misc"": """ & device.misc & """"
  response.write "}"

  'if dev.status_support then response.write ", supports status report"
  'if (dev.misc and &H1) <> 0 then response.write ", preset dim"
  'if (dev.misc and &H2) <> 0 then response.write ", extended dim"
  'if (dev.misc and &H4) <> 0 then response.write ", SmartLinc"
  'if (dev.misc and &H8) <> 0 then response.write ", no logging"
  'if (dev.misc and &H10) <> 0 then response.write ", status-only"
  'if (dev.misc and &H20) <> 0 then response.write ", hidden"
  'if (dev.misc and &H40) <> 0 then response.write ", thermostat"
  'if (dev.misc and &H80) <> 0 then response.write ", included in power-fail"
  'if (dev.misc and &H100) <> 0 then response.write ", show values"
  'if (dev.misc and &H200) <> 0 then response.write ", auto voice command"
  'if (dev.misc and &H400) <> 0 then response.write ", confirm voice command"
  'if (dev.misc and &H800) <> 0 then response.write ", Compose device"
  'if (dev.misc and &H1000) <> 0 then response.write ", Z-Wave"
  'if (dev.misc and &H2000) <> 0 then response.write ", other direct-level dim"
  'if (dev.misc and &H4000) <> 0 then response.write ", plugin status call"
  'if (dev.misc and &H8000) <> 0 then response.write ", plugin value call"

end sub

' *****************************************************************************
sub DeviceOn(id)
  ' find the device validating if it exists
  if hs.DeviceExists(id) = -1 then
      ShowError("No device exists with id " & id & ".")
      exit sub
    end if

  set dev = hs.GetDeviceByRef(hs.DeviceExistsRef(id))

  if dev.can_dim then
    hs.ExecX10 id,"ddim",100
    ' TODO make sure we get an updated device json object back
    ' currently it is the old one, with the old value / status
    ShowDevice(id)
  else ' just turn it on
    hs.ExecX10 id,"on"
    ' TODO make sure we get an updated device json object back
    ' currently it is the old one, with the old value / status
    ShowDevice(id)
  end if

  hs.WriteLog "JSONAPI","Device action: " & "Turned on" & " the " & dev.location & " " & dev.name & " at " & id & " from " & hs.WebLoggedInUser & " at " & request.ServerVariables("REMOTE_ADDR")
end sub

' *****************************************************************************
sub DeviceOff(id)
  ' find the device validating if it exists
  if hs.DeviceExists(id) = -1 then
      ShowError("No device exists with id " & id & ".")
      exit sub
    end if

  set dev = hs.GetDeviceByRef(hs.DeviceExistsRef(id))

  hs.ExecX10 id,"off"
  ' TODO make sure we get an updated device json object back
  ' currently it is the old one, with the old value / status
  ShowDevice(id)

  hs.WriteLog "JSONAPI","Device action: " & "Turned off" & " the " & dev.location & " " & dev.name & " at " & id & " from " & hs.WebLoggedInUser & " at " & request.ServerVariables("REMOTE_ADDR")

end sub


' *****************************************************************************
sub SetDeviceValue(id, value)
  ' find the device validating if it exists
  if hs.DeviceExists(id) = -1 then
      ShowError("No device exists with id " & id & ".")
      exit sub
    end if

  set dev = hs.GetDeviceByRef(hs.DeviceExistsRef(id))

  ' what's the current and target dim level?
  l = hs.DeviceValue(id)
  k = value
  if k < 0 or k > 100 then
      response.write "Invalid dim level " & k & " ignored."
      exit sub
    end if

  if dev.can_dim then
    hs.ExecX10 id,"ddim",k
    ' TODO make sure we get an updated device json object back
    ' currently it is the old one, with the old value / status
    ShowDevice(id)
  else ' just turn it on
    ShowError("Device not dimmable")
  end if

  hs.WriteLog "JSONAPI","Device action: " & "Dimmed" & " the " & dev.location & " " & dev.name & " at " & id & " from " & hs.WebLoggedInUser & " at " & request.ServerVariables("REMOTE_ADDR")

end sub


' *****************************************************************************
sub ShowEvents()
  dim events

  set events = hs.GetEventEnumerator

  response.write "["

  ' loop through all devices
  if IsObject(events) then
    do while not events.Finished
      set evt = events.GetNext
      if not evt is nothing then
        ' print event
        ShowEvent(evt)
        response.write ", "
      end if
    loop
  end if

  response.write "]"
end sub


' *****************************************************************************
sub ShowEvent(evt)
  'Render an event
  response.write "{"
  response.write """id"": """ & EncodeString(evt.name) & """, "
  response.write """name"": """ & evt.name & """, "
  response.write """group"": """ & evt.group & """"
  response.write "}"
end sub


sub RunEvent(id)
    hs.TriggerEvent Replace(id, "_", " ")
    response.write "{""status"": ""Event " & Replace(id, "_", " ") & " triggered""}"
    hs.WriteLog "JSONAPI","Triggering event " & Replace(id, "_", " ") & " from " & hs.WebLoggedInUser & " at " & request.ServerVariables("REMOTE_ADDR")
end sub


' *****************************************************************************
function ShowError(msg)
  response.write "{""error"": """ & msg & """}"
end function


' *****************************************************************************
sub BuildLocationsList()
  ' Build a list of locations for later use (up to 64 locations)
  dim dev, devices, thisloc
  numlocations = 0

  ' loop through devices
  set devices = hs.GetDeviceEnumerator
  if IsObject(devices) then
      do while not devices.Finished
        set dev = devices.GetNext
        if not dev is nothing then
            if opt_showhidden = "yes" or (dev.misc and &H20) = 0 then
                l = 1 ' yes, add it
                thisloc = dev.location
                if thisloc = "" then thisloc = "No Location"
                ' see if this location is already listed
                for k = 1 to numlocations
                  if lcase(thisloc) = lcase(locations(k)) then
                      l = 0 ' don't bother, it's already there
                      exit for
                    end if
                   next
                ' if we're due to add it to the list,
                if l = 1 and numlocations < 64 then ' add it
                    numlocations = numlocations + 1
                    locations(numlocations) = thisloc
                  end if
              end if
          end if
        loop
    end if

  ' sort the list -- simple selection sort
  if numlocations > 2 then ' sort the list
      for i = 1 to numlocations-1
        k = i
        for j = i+1 to numlocations
          if locations(j) < locations(k) then k = j
          next
        if k <> i then
            s = locations(k)
            locations(k) = locations(i)
            locations(i) = s
          end if
        next
    end if
end sub



' *****************************************************************************
function EncodeString(s)
  ' converts a string to URL-encoding (e.g., Frank%26s_Room)
  dim i, c, result
  result = ""
  for i = 1 to len(s)
    c = mid(s,i,1)
    if c = " " then
        result = result & "_"
      elseif c = "." or c = "-" then
        result = result & c
      elseif c < "0" or (c > "9" and c < "A") or (c > "Z" and c < "a") or (c > "z") then
        result = result & "%" & hex(asc(c))
      else
        result = result & c
      end if
    next
  EncodeString = result
end function


function DecodeString(s)
  ' reverses the above
  dim i, c, result
  result = ""
  for i = 1 to len(s)
    c = mid(s,i,1)
    if c = "%" then
        result = result & chr(mid(s,i+1,2))
        i = i + 2
      else
        result = result & c
      end if
    next
  DecodeString = result
end function


' *****************************************************************************
function GetUserAuthorizations(username)
  dim sUserList, sUsers, sUser, sUsername, sRights, i
  sUserList = hs.GetUsers
  sUsers = Split(sUserList,",")
  for i = 0 to UBound(sUsers)
    sUser = sUsers(i)
    sUsername = left(sUser,instr(sUser,"|")-1)
    sRights = cint(trim(mid(sUser,instr(sUser,"|")+1)))
    if sUsername = username then
        GetUserAuthorizations = sRights
        exit function
      end if
    next
  GetUserAuthorizations = 0
end function



' *****************************************************************************
sub hs_install()
  ' nothing to do
end sub

sub hs_uninstall()
  ' nothing to do
end sub


' *****************************************************************************
function titleCase(phrase) 
  dim words, i 
  words = split(phrase," ") 
  for i = 0 to ubound(words) 
    words(i) = ucase(left(words(i),1)) & lcase(mid(words(i),2)) 
  next 
  titleCase = join(words," ") 
end function 

%>
