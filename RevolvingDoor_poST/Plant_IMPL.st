IF NOT initialized THEN
    GVL.doorAngle            := 0.0;
    GVL.presenceSideA        := FALSE;
    GVL.presenceSideB        := FALSE;
    GVL.presenceTriggerA     := FALSE;
    GVL.presenceTriggerB     := FALSE;
    presenceTimerA           := 0;
    presenceTimerB           := 0;
    initialized              := TRUE;
END_IF

(* Presence hold - trigger is set by visualization, timer keeps presence active *)
IF GVL.presenceTriggerA THEN
    presenceTimerA := PRESENCE_HOLD;
    GVL.presenceTriggerA := FALSE;
END_IF
GVL.presenceSideA := (presenceTimerA > 0);
IF presenceTimerA > 0 THEN
    presenceTimerA := presenceTimerA - 1;
END_IF

IF GVL.presenceTriggerB THEN
    presenceTimerB := PRESENCE_HOLD;
    GVL.presenceTriggerB := FALSE;
END_IF
GVL.presenceSideB := (presenceTimerB > 0);
IF presenceTimerB > 0 THEN
    presenceTimerB := presenceTimerB - 1;
END_IF

(* MotorSim *)
GVL.isRotating := GVL.motorOn;

(* DoorRotationSim *)
IF GVL.isRotating THEN
    IF GVL.rotationForward THEN
        GVL.doorAngle := GVL.doorAngle - ROTATION_SPEED;
    ELSE
        GVL.doorAngle := GVL.doorAngle + ROTATION_SPEED;
    END_IF
END_IF

IF GVL.doorAngle >= MAX_ANGLE THEN
    GVL.doorAngle := GVL.doorAngle - MAX_ANGLE;
END_IF
IF GVL.doorAngle < 0.0 THEN
    GVL.doorAngle := GVL.doorAngle + MAX_ANGLE;
END_IF