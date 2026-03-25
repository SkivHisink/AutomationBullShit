{attribute 'qualified_only' := 'false'}
VAR_GLOBAL
    (* Presence sensors *)
    presenceSideA           : BOOL;
    presenceSideB           : BOOL;
    (* Presence triggers - visualization writes here *)
    presenceTriggerA        : BOOL;
    presenceTriggerB        : BOOL;
    (* Partition pressure sensor *)
    partitionPressure       : BOOL;
    (* Manual rotation buttons *)
    rotateForwardButton     : BOOL;
    rotateBackwardButton    : BOOL;
    (* Rotation direction: FALSE=counter-clockwise (default), TRUE=clockwise *)
    rotationClockwise       : BOOL;
    (* Door angle sensor *)
    doorAngle               : REAL;
(* Rotation status *)
    isRotating              : BOOL;
    (* Motor control *)
    motorOn                 : BOOL;
    rotationForward         : BOOL;
    (* Motor enable switch *)
    motorEnable             : BOOL := TRUE;
    (* Emergency stop *)
    emergencyStop           : BOOL;
    (* Traffic lights *)
    lightSideA              : BOOL;
    lightSideB              : BOOL;
END_VAR
