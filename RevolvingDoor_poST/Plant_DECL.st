PROGRAM Plant
VAR CONSTANT
    ROTATION_SPEED  : REAL := 0.1;
    MAX_ANGLE       : REAL := 360.0;
    PRESENCE_HOLD   : INT := 500;
END_VAR
VAR
    initialized      : BOOL := FALSE;
    presenceTimerA   : INT := 0;
    presenceTimerB   : INT := 0;
END_VAR
