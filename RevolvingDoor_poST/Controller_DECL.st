PROGRAM Controller
VAR CONSTANT
    ROTATION_TIME   : INT := 100;
    REVERSE_TICKS   : INT := 50;    (* 5 degrees at ROTATION_SPEED=0.1 *)
END_VAR
VAR
    (* MainControl state: 0=idle, 1=rotating, 2=manualRotation, 3=pressure *)
    mainState               : INT := 0;
    timer                   : INT := 0;
    manualActive            : BOOL := FALSE;
    manualDirectionForward  : BOOL := TRUE;
    (* Pressure sub-process: 0=released, 1=reversing, 2=stopped *)
    pauseState              : INT := 0;
    reverseCounter          : INT := 0;
END_VAR
