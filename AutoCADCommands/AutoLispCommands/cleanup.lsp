(defun c:CLEANUP (/ oldCMDECHO)
  ;; Remember current command echo setting, then turn it off:
  (setq oldCMDECHO (getvar "CMDECHO"))
  (setvar "CMDECHO" 0)

  ;; Use SETBYLAYER on everything in the current space
  ;;   "Y"  => Change ByBlock to ByLayer?
  ;;   "Y"  => Include blocks?
  (command
    "_.SETBYLAYER" "All" "" "Y" "Y")

  ;; Purge the entire drawing; "All", "*" (everything), "No" to confirm each
  (command "_.PURGE" "All" "*" "N")

  ;; Audit the drawing; "Yes" to fix errors
  (command "_.AUDIT" "Y")

  ;; Restore command echo setting:
  (setvar "CMDECHO" oldCMDECHO)
  (princ)
)
