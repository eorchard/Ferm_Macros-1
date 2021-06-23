'#####################################################################
' script parameters
'#####################################################################
 
'### These parameters get changed every run ###
  dim DO_low_trigger as double = 30                 '[%] usually DO setpoint
  dim DO_high_trigger as double = 40                '[%] feed trigger
  dim bolus_size as double = 30                     '[g/L]
  dim Wait_Before_DO_Detection as double = 15/60    '[h] blind time time before looking for high DO trigger after hitting low trigger

'### These parameters should remain static generally ###
  dim Wait_Before_Feed_Start as double = 1/60       '[h] Time before high trigger and feeding should be 0.033 for 2min delay - to prevent false alarms
  dim Feed_Rate as double = 200                     '[ml/h] default 200mL/hr
  dim Feed_Duration as double                       '[h] feed duration to be calculated during script
  dim Feed_Duration_timebased as double = 2/60      '[h] for use in time based feeding rather than bolus
  dim feed_density as double = 1008                 '[g/L], typically 1008 for 80% glycerol
  dim Minimum_Slope_For_Feed = 5                    'prevent feed shots from creeping DO
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
          'Reset totalizer values
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

            'Identify Citric Acid spike if DO > Low Trigger and pH > 7
            if (.ExtA < 1) And (.PHPV > 7) then
                .phase = .phase - 4

                'Turn on Pump A if needed
                .PumpAActive = 1

                .LogMessage("Turning on Pump A following Citric Acid spike")
                .LogMessage("Entering phase: Waiting for DO falling under " & DO_low_trigger & "%")

            'Start slope calculation once DO > Low Trigger
            else if (.ExtA >= 1) then
                .phase = .phase + 1
                .LogMessage("Entering phase: Approaching DO high trigger of " & DO_high_trigger &"%")
            end if
        end if

      case 5
        'Slope calculation
        DO_Slope = ((DO_high_trigger/100)-(DO_low_trigger/100))/(.Runtime_H - .PhaseStart_H)

        if .DOPV > DO_high_trigger then

          'Check for minimum slope 
          if ((DO_Slope > Minimum_Slope_For_Feed) Or (.ExtA < 1)) then
            .phase = .phase + 1
            .LogMessage("Entering phase: Waiting for high DO longer than " & Wait_Before_Feed_Start & "h")

          'Slope is below the minimum, don't feed
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
            .phase = .phase + 1
            .LogMessage("Entering phase: Starting feeds A")
            .FASP = Feed_Rate

            'Count number of DO spikes, currently used to identify Citric Acid spike but may be useful in a future script
            .ExtA = .ExtA + 1     
        end if

      case 7
        Feed_Duration = ((.VPV * bolus_size)/feed_density)/Feed_Rate
        if .Runtime_H - .PhaseStart_H > Feed_Duration then
          .phase = .phase - 5
          .LogMessage("Feeds Complete, Waiting for DO falling under " & DO_low_trigger & "%")
          .FASP = 0
        end if
'       if .Runtime_H - .PhaseStart_H > Feed_Duration_timebased then
'         .phase = .phase - 5
'         .LogMessage("Feeds Complete, Waiting for DO falling under " & DO_low_trigger & "%")
'         .FASP = 0
'       end if
    end select
   
    
  end with
end if
