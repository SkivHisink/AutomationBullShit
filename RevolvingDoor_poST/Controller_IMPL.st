(* Emergency stop - disable motor *)
IF GVL.emergencyStop THEN
    GVL.motorEnable := FALSE;
END_IF

(* Motor disabled - immediate stop, reset to idle *)
IF NOT GVL.motorEnable THEN
    GVL.motorOn := FALSE;
    GVL.rotationForward := NOT GVL.rotationClockwise;
    GVL.lightSideA := FALSE;
    GVL.lightSideB := FALSE;
	GVL.presenceSideA := FALSE;	
    GVL.presenceSideB := FALSE;
    manualActive := FALSE;
    timer := 0;
    pauseState := 0;
    mainState := 0;
ELSE

CASE mainState OF

0: (* idle *)
    GVL.lightSideA := TRUE;
    GVL.lightSideB := TRUE;
    IF GVL.rotateForwardButton AND NOT GVL.rotateBackwardButton AND NOT GVL.partitionPressure THEN
        manualActive := TRUE;
        manualDirectionForward := TRUE;
        GVL.motorOn := TRUE;
        GVL.rotationForward := NOT GVL.rotationClockwise;
        GVL.lightSideA := TRUE;
        GVL.lightSideB := TRUE;
        mainState := 2;

    ELSIF GVL.rotateBackwardButton AND NOT GVL.rotateForwardButton AND NOT GVL.partitionPressure THEN
        manualActive := TRUE;
        manualDirectionForward := FALSE;
        GVL.motorOn := TRUE;
        GVL.rotationForward := GVL.rotationClockwise;
        GVL.lightSideA := TRUE;
        GVL.lightSideB := TRUE;
        mainState := 2;

    ELSIF (GVL.presenceSideA OR GVL.presenceSideB) AND NOT GVL.partitionPressure THEN
        manualActive := FALSE;
        manualDirectionForward := TRUE;
        GVL.motorOn := TRUE;
        GVL.rotationForward := NOT GVL.rotationClockwise;
        GVL.lightSideA := TRUE;
        GVL.lightSideB := TRUE;
        timer := ROTATION_TIME;
        mainState := 1;
    END_IF

1: (* rotating *)
    IF (GVL.presenceSideA OR GVL.presenceSideB) THEN
        timer := ROTATION_TIME;
    END_IF

    IF GVL.partitionPressure THEN
        GVL.motorOn := FALSE;
        GVL.rotationForward := NOT GVL.rotationForward;
        GVL.lightSideA := FALSE;
        GVL.lightSideB := FALSE;
        reverseCounter := REVERSE_TICKS;
        pauseState := 1;
        mainState := 3;
    ELSIF NOT (GVL.presenceSideA OR GVL.presenceSideB) THEN
        timer := timer - 1;
    END_IF

    IF timer <= 0 AND mainState = 1 THEN
        GVL.motorOn := FALSE;
        GVL.rotationForward := NOT GVL.rotationClockwise;
        GVL.lightSideA := TRUE;
        GVL.lightSideB := TRUE;
        mainState := 0;
    END_IF

2: (* manualRotation *)
    IF GVL.partitionPressure THEN
        GVL.motorOn := FALSE;
        GVL.rotationForward := NOT GVL.rotationForward;
        GVL.lightSideA := FALSE;
        GVL.lightSideB := FALSE;
        reverseCounter := REVERSE_TICKS;
        pauseState := 1;
        mainState := 3;

    ELSIF GVL.rotateForwardButton AND GVL.rotateBackwardButton THEN
        GVL.motorOn := FALSE;
        GVL.rotationForward := NOT GVL.rotationClockwise;
        GVL.lightSideA := TRUE;
        GVL.lightSideB := TRUE;
        manualActive := FALSE;
        mainState := 0;

    ELSIF manualDirectionForward AND GVL.rotateForwardButton THEN
        GVL.motorOn := TRUE;
        GVL.rotationForward := NOT GVL.rotationClockwise;
        GVL.lightSideA := TRUE;
        GVL.lightSideB := TRUE;

    ELSIF (NOT manualDirectionForward) AND GVL.rotateBackwardButton THEN
        GVL.motorOn := TRUE;
        GVL.rotationForward := GVL.rotationClockwise;
        GVL.lightSideA := TRUE;
        GVL.lightSideB := TRUE;

    ELSE
        GVL.motorOn := FALSE;
        GVL.rotationForward := NOT GVL.rotationClockwise;
        GVL.lightSideA := TRUE;
        GVL.lightSideB := TRUE;
        manualActive := FALSE;
        mainState := 0;
    END_IF

3: (* pressure reaction *)
    IF pauseState = 1 THEN
        (* Reversing 5 degrees *)
        GVL.motorOn := TRUE;
        reverseCounter := reverseCounter - 1;
        IF reverseCounter <= 0 THEN
            GVL.motorOn := FALSE;
            pauseState := 2;
        END_IF
    END_IF

    IF pauseState = 2 THEN
        (* Stopped until pressure released *)
        GVL.motorOn := FALSE;
        IF NOT GVL.partitionPressure THEN
            pauseState := 0;
        END_IF
    END_IF

    IF pauseState = 0 THEN
        (* Pressure released - resume *)
        IF manualActive THEN
            IF manualDirectionForward AND GVL.rotateForwardButton THEN
                GVL.motorOn := TRUE;
                GVL.rotationForward := NOT GVL.rotationClockwise;
                GVL.lightSideA := TRUE;
                GVL.lightSideB := TRUE;
                mainState := 2;
            ELSIF (NOT manualDirectionForward) AND GVL.rotateBackwardButton THEN
                GVL.motorOn := TRUE;
                GVL.rotationForward := GVL.rotationClockwise;
                GVL.lightSideA := TRUE;
                GVL.lightSideB := TRUE;
                mainState := 2;
            ELSE
                GVL.motorOn := FALSE;
                GVL.rotationForward := NOT GVL.rotationClockwise;
                GVL.lightSideA := TRUE;
                GVL.lightSideB := TRUE;
                manualActive := FALSE;
                mainState := 0;
            END_IF
        ELSIF timer > 0 THEN
            GVL.motorOn := TRUE;
            GVL.rotationForward := NOT GVL.rotationClockwise;
            GVL.lightSideA := TRUE;
            GVL.lightSideB := TRUE;
            mainState := 1;
        ELSE
            GVL.motorOn := FALSE;
            GVL.rotationForward := NOT GVL.rotationClockwise;
            GVL.lightSideA := TRUE;
            GVL.lightSideB := TRUE;
            mainState := 0;
        END_IF
    END_IF

END_CASE

END_IF (* motorEnable / emergencyStop *)
