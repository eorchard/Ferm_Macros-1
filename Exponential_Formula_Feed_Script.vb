'#####################################################################
' script parameters
'#####################################################################
 
'### These parameters get changed every run ###
  dim DO_low_trigger as double = 30                                       '[%] usually DO setpoint
  dim DO_high_trigger as double = 40                                      '[%] feed trigger
  dim Exponential_Feed_Duration as double = 8                             '# of hours for exponential feed
  dim Exponential_Feed_Initial_Rate as double = 200                       'initial feed rate
  dim Linear_Feed_Rate as double = 200                                    'linear feed after exponential feed duration
  dim Specific_Growth_Rate = 5                                            'specific growth rate from shake flask data
  dim Wait_Before_DO_Detection as double = 15/60                          '[h] blind time time before looking for high DO trigger after hitting low trigger

'### These parameters should remain static generally ###
  dim Wait_Before_Feed_Start as double = 1/60                            '[h] Time before high trigger and feeding should be 0.033 for 2min delay - to prevent false alarms
  dim Exponential_Feed_Rate as double = Exponential_Feed_Initial_Rate * (2.72^(Specific_Growth_Rate*Exponential_Feed_Duration))
 
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
          .LogMessage("Entering phase: Starting exponential feed)
          .FASP = Exponential_Feed_Rate
            if .Runtime_H - .PhaseStart_H > Exponential_Feed_Duration then
              .phase = .phase + 1
            end if
        end if

      case 6
          .LogMessage("Entering phase: Starting linear feed)
          .FASP = Linear_Feed_Rate
 
    end select
 
  end with
end if
 
