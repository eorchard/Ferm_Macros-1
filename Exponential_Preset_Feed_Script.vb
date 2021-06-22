'#####################################################################
' script parameters
'#####################################################################
 
'### These parameters get changed every run ###
  dim DO_low_trigger as double = 30                                       '[%] usually DO setpoint
  dim DO_high_trigger as double = 40                                      '[%] feed trigger
  dim Exponential_Feed_Steps as double = 15                      'number of exponential feed phases
  dim Feed_Rate_Step_1 as double = 200                              '[ml/h] default 200mL/hr
  dim Feed_Rate_Step_2 as double = 200                              '[ml/h] default 200mL/hr
  dim Feed_Rate_Step_3 as double = 200                              '[ml/h] default 200mL/hr
  dim Feed_Rate_Step_4 as double = 200                              '[ml/h] default 200mL/hr
  dim Feed_Rate_Step_5 as double = 200                              '[ml/h] default 200mL/hr
  dim Feed_Rate_Step_6 as double = 200                              '[ml/h] default 200mL/hr
  dim Feed_Rate_Step_7 as double = 200                              '[ml/h] default 200mL/hr
  dim Feed_Rate_Step_8 as double = 200                              '[ml/h] default 200mL/hr
  dim Feed_Rate_Step_9 as double = 200                              '[ml/h] default 200mL/hr
  dim Feed_Rate_Step_10 as double = 200                            '[ml/h] default 200mL/hr
  dim Feed_Rate_Step_11 as double = 200                            '[ml/h] default 200mL/hr
  dim Feed_Rate_Step_12 as double = 200                            '[ml/h] default 200mL/hr
  dim Feed_Rate_Step_13 as double = 200                            '[ml/h] default 200mL/hr
  dim Feed_Rate_Step_14 as double = 200                            '[ml/h] default 200mL/hr
  dim Feed_Rate_Step_15 as double = 200                            '[ml/h] default 200mL/hr
  dim EFT_For_Step_2 as double = 8                                        'EFT for step up in exponential feed
  dim EFT_For_Step_3 as double = 8                                        'EFT for step up in exponential feed
  dim EFT_For_Step_4 as double = 8                                        'EFT for step up in exponential feed
  dim EFT_For_Step_5 as double = 8                                        'EFT for step up in exponential feed
  dim EFT_For_Step_6 as double = 8                                        'EFT for step up in exponential feed
  dim EFT_For_Step_7 as double = 8                                        'EFT for step up in exponential feed
  dim EFT_For_Step_8 as double = 8                                        'EFT for step up in exponential feed
  dim EFT_For_Step_9 as double = 8                                        'EFT for step up in exponential feed
  dim EFT_For_Step_10 as double = 8                    'EFT for step up in exponential feed
  dim EFT_For_Step_11 as double = 8                    'EFT for step up in exponential feed
  dim EFT_For_Step_12 as double = 8                    'EFT for step up in exponential feed
  dim EFT_For_Step_13 as double = 8                    'EFT for step up in exponential feed
  dim EFT_For_Step_14 as double = 8                    'EFT for step up in exponential feed
  dim EFT_For_Step_15 as double = 8                    'EFT for step up in exponential feed
  dim Wait_Before_DO_Detection as double = 15/60        '[h] blind time time before looking for high DO trigger after hitting low trigger

'### These parameters should remain static generally ###
  dim Wait_Before_Feed_Start as double = 1/60                 '[h] Time before high trigger and feeding should be 0.033 for 2min delay - to prevent false alarms
 
'#####################################################################
 
 
if p isnot nothing then
  with p 
 
    select case .phase
      case 0 
        'init
        .phase = .phase + 1
        .LogMessage("Entering phase: Waiting for InoculationTime start")

      case 1
        if .InoculationStart_H > 0 then
          .phase = .phase + 1
          .LogMessage("Entering phase: Waiting for DO falling under " & DO_low_trigger & "%")
        end if

      case 2
        if .DOPV < DO_low_trigger then
          .phase = .phase + 1
          .LogMessage("Entering phase: Waiting blind time before DO rising detection for " & Wait_Before_DO_Detection & "h")
        end if

      case 3
        if  .Runtime_H - .PhaseStart_H > Wait_Before_DO_Detection then
          .phase = .phase + 1
          .LogMessage("Entering phase: Waiting for DO > "& DO_high_trigger &"%")
        end if

      case 4
        if .DOPV > DO_high_trigger then
          .phase = .phase + 1
          .LogMessage("Entering phase: Waiting for high DO longer than " & Wait_Before_Feed_Start & "h")
        end if

      case 5
        if .DOPV < DO_high_trigger then
          .phase = .phase - 1
          .LogMessage("Entering phase: Waiting for DO > "& DO_high_trigger &"%")
        end if 
        if .Runtime_H - .PhaseStart_H > Wait_Before_Feed_Start then
          .LogMessage("Entering phase: Starting exponential feed step 1 of " & Exponential_Feed_Steps)
          .FASP = Feed_Rate_Step_1
                  if Exponential_Feed_Steps > 1 then
                  .phase = .phase + 1
                  end if

      case 6
            if .Runtime_H > EFT_For_Step_2 then
              .LogMessage("Entering phase: Starting exponential feed step 2 of " & Exponential_Feed_Steps)
              .FASP = Feed_Rate_Step_2
                      if Exponential_Feed_Steps > 2 then
                      .phase = .phase + 1
                      end if
            end if   

      case 7
            if .Runtime_H > EFT_For_Step_3 then
          .LogMessage("Entering phase: Starting exponential feed step 3 of " & Exponential_Feed_Steps)
          .FASP = Feed_Rate_Step_3
                  if Exponential_Feed_Steps > 3 then
                  .phase = .phase + 1
                  end if
        end if   

      case 8
                if .Runtime_H > EFT_For_Step_4 then
          .LogMessage("Entering phase: Starting exponential feed step 4 of " & Exponential_Feed_Steps)
          .FASP = Feed_Rate_Step_4
                  if Exponential_Feed_Steps > 4 then
                  .phase = .phase + 1
                  end if
        end if 

      case 9
                if .Runtime_H > EFT_For_Step_5 then
          .LogMessage("Entering phase: Starting exponential feed step 5 of " & Exponential_Feed_Steps)
          .FASP = Feed_Rate_Step_5
                  if Exponential_Feed_Steps > 5 then
                  .phase = .phase + 1
                  end if
        end if 

      case 10
                if .Runtime_H > EFT_For_Step_6 then
          .LogMessage("Entering phase: Starting exponential feed step 6 of " & Exponential_Feed_Steps)
          .FASP = Feed_Rate_Step_6
                  if Exponential_Feed_Steps > 6 then
                  .phase = .phase + 1
                  end if
        end if 

      case 11
                if .Runtime_H > EFT_For_Step_7 then
          .LogMessage("Entering phase: Starting exponential feed step 7 of " & Exponential_Feed_Steps)
          .FASP = Feed_Rate_Step_7
                  if Exponential_Feed_Steps > 7 then
                  .phase = .phase + 1
                  end if
        end if 

      case 12
                if .Runtime_H > EFT_For_Step_8 then
          .LogMessage("Entering phase: Starting exponential feed step 8 of " & Exponential_Feed_Steps)
          .FASP = Feed_Rate_Step_8
                  if Exponential_Feed_Steps > 8 then
                  .phase = .phase + 1
                  end if
        end if 

      case 13
                if .Runtime_H > EFT_For_Step_9 then
          .LogMessage("Entering phase: Starting exponential feed step 9 of " & Exponential_Feed_Steps)
          .FASP = Feed_Rate_Step_9
                  if Exponential_Feed_Steps > 9 then
                  .phase = .phase + 1
                  end if
        end if 

      case 14
                if .Runtime_H > EFT_For_Step_10 then
          .LogMessage("Entering phase: Starting exponential feed step 10 of " & Exponential_Feed_Steps)
          .FASP = Feed_Rate_Step_10
                  if Exponential_Feed_Steps > 10 then
                  .phase = .phase + 1
                  end if
        end if 

      case 15
                if .Runtime_H > EFT_For_Step_11 then
          .LogMessage("Entering phase: Starting exponential feed step 11 of " & Exponential_Feed_Steps)
          .FASP = Feed_Rate_Step_11
                  if Exponential_Feed_Steps > 11 then
                  .phase = .phase + 1
                  end if
        end if 

      case 16
                if .Runtime_H > EFT_For_Step_12 then
          .LogMessage("Entering phase: Starting exponential feed step 12 of " & Exponential_Feed_Steps)
          .FASP = Feed_Rate_Step_12
                  if Exponential_Feed_Steps > 12 then
                  .phase = .phase + 1
                  end if
        end if 

      case 17
                if .Runtime_H > EFT_For_Step_13 then
          .LogMessage("Entering phase: Starting exponential feed step 13 of " & Exponential_Feed_Steps)
          .FASP = Feed_Rate_Step_13
                  if Exponential_Feed_Steps > 13 then
                  .phase = .phase + 1
                  end if
        end if 

      case 18
                if .Runtime_H > EFT_For_Step_14 then
          .LogMessage("Entering phase: Starting exponential feed step 14 of " & Exponential_Feed_Steps)
          .FASP = Feed_Rate_Step_14
                  if Exponential_Feed_Steps > 14 then
                  .phase = .phase + 1
                  end if
        end if 

      case 19
                if .Runtime_H > EFT_For_Step_15 then
          .LogMessage("Entering phase: Starting exponential feed step 15 of " & Exponential_Feed_Steps)
          .FASP = Feed_Rate_Step_15
        end if  
    end select
 
  end with
end if
