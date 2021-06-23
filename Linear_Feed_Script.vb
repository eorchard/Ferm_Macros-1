'#####################################################################
' script parameters
'#####################################################################
 
'### These parameters get changed every run ###
  dim DO_low_trigger as double = 30                 '[%] usually DO setpoint
  dim DO_high_trigger as double = 40                '[%] feed trigger
  dim Linear_Feed_Duration as double = 8            '# of hours for linear feed
  dim Linear_Feed_Initial_Rate as double = 10           'initial feed rate
  dim Linear_Feed_Final_Rate as double = 100            'linear feed will climb to this rate
  dim Linear_Feed_Constant_Rate as double = 80         'linear feed will change to this rate after climbing to final rate
  dim Wait_Before_DO_Detection as double = 1/60        '[h] blind time time before looking for high DO trigger after hitting low trigger

'### These parameters should remain static generally ###
  dim Wait_Before_Feed_Start as double = 1/60       '[h] Time before high trigger and feeding should be 0.033 for 2min delay - to prevent false alarms
  dim Linear_Feed_Rate as double = Linear_Feed_Initial_Rate + ((Linear_Feed_Final_Rate - Linear_Feed_Initial_Rate)/Linear_Feed_Duration)*(p.Runtime_H - p.PhaseStart_H)
  dim Minimum_Slope_For_Feed as Integer = 5         'prevent feed shots from creeping DO
  dim DO_Slope as double                            'slope will be calculated during script
 
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
          .SetVAPV(0)
          .SetVBPV(0)
          .phase = .phase + 1
          .LogMessage("Entering phase: Totalizer values reset. Waiting for DO falling under " & DO_low_trigger & "%")
        end if

      case 2
        if .DOPV < DO_low_trigger then
          .phase = .phase + 1
          .LogMessage("Entering phase: Waiting blind time before DO rising detection for " & Wait_Before_DO_Detection & "h")
        end if

      case 3
         if .Runtime_H - .PhaseStart_H > Wait_Before_DO_Detection then
          .phase = .phase + 1
          .LogMessage("Entering phase: Waiting for DO > "& DO_high_trigger &"%")
        end if

      case 4
       if .DOPV > DO_low_trigger then
          .phase = .phase + 1
          .LogMessage("Entering phase: Approaching DO high trigger of " & DO_high_trigger &"%")
        end if

      case 5
        'Slope calculation
        DO_Slope = ((DO_high_trigger/100)-(DO_low_trigger/100))/(.Runtime_H - .PhaseStart_H)
        if .DOPV > DO_high_trigger then

          'Check for minimum slope 
          if ((DO_Slope > Minimum_Slope_For_Feed) Or (.ExtA < 1))
            .phase = .phase + 1
            .LogMessage("Entering phase: Waiting for high DO longer than " & Wait_Before_Feed_Start & "h")

          else
            .phase = .phase - 3
            .LogMessage("Slope of " & DO_Slope & " doesn't meet minimum slope of " & Minimum_Slope_For_Feed)
            .LogMessage("Entering phase: Waiting for DO falling under " & DO_low_trigger & "%")
          end if
        end if

        'Slope falls back below low trigger
        if .DOPV < DO_low_trigger then
          .phase = .phase - 2
          .LogMessage("Entering phase: Waiting for DO > "& DO_high_trigger &"%")
        end if

      case 6
        if .DOPV < DO_high_trigger then
          .phase = .phase - 1
          .LogMessage("Entering phase: Waiting for DO > "& DO_high_trigger &"%")
        end if 
        if .Runtime_H - .PhaseStart_H > Wait_Before_Feed_Start then
          if .ExtA < 1 then
            .phase = .phase - 4
            .PumpAActive = 1
           .LogMessage("Turning on Pump A following Citric Acid spike")
            .LogMessage("Entering phase: Waiting for DO falling under " & DO_low_trigger & "%")
          else
            .LogMessage("Entering phase: Starting linear feed")
            .phase = .phase + 1
          end if
          .ExtA = .ExtA + 1     'this indicates number of DO spikes, currently used to identify Citric Acid spike but may be useful in a future script
        end if

      case 7
        .FASP = Linear_Feed_Rate
                if .Runtime_H - .PhaseStart_H > Linear_Feed_Duration then
          .LogMessage("Entering phase: Starting constant linear feed")
          .phase = .phase + 1
                end if

      case 8
        .FASP = Linear_Feed_Constant_Rate  
 
    end select
 
  end with
end if
