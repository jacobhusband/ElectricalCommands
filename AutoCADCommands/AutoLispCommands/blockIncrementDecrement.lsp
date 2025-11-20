;; BlockAttributeTools.lsp
;;
;; A suite of tools to increment, decrement, or set the first numeric attribute of selected block references.
;;
;; To Use:
;; 1. Load this file into AutoCAD using the APPLOAD command.
;; 2. Select one or more block references in your drawing.
;; 3. Type one of the following commands and press Enter:
;;
;;    C1: Increments the value by 1.
;;    D1: Decrements the value by 1.
;;    C2: Prompts for a value to INCREMENT by.
;;    D2: Prompts for a value to DECREMENT by.
;;    N1: Prompts for a new value to SET the attribute to.
;;
;; Note: If no blocks are pre-selected, the commands will prompt you to select them.

;; =============================================================================
;; == COMMAND C1: Increment by 1 (Original Script)
;; =============================================================================
(defun c:C1 (/ *error* old_cmdecho sel i count ent att_ent att_edata old_val new_val found_att)
  (defun *error* (msg)
    (if old_cmdecho (setvar "CMDECHO" old_cmdecho))
    (if (not (member msg '("Function cancelled" "quit / exit abort"))) (princ (strcat "\nError: " msg)))
    (princ)
  )
  (setq old_cmdecho (getvar "CMDECHO"))
  (setvar "CMDECHO" 0)
  (princ "\nSelect block(s) with attributes to increment by 1: ")
  (setq sel (ssget '((0 . "INSERT") (66 . 1))))
  (if sel
    (progn
      (setq i 0 count 0)
      (repeat (sslength sel)
        (setq ent (ssname sel i) att_ent (entnext ent) found_att nil)
        (while (and att_ent (not found_att))
          (setq att_edata (entget att_ent))
          (if (= "ATTRIB" (cdr (assoc 0 att_edata)))
            (progn
              (setq old_val (cdr (assoc 1 att_edata)))
              (if (and old_val (numberp (read old_val))) (setq found_att T))
            )
          )
          (if (not found_att) (setq att_ent (entnext att_ent)))
        )
        (if found_att
          (progn
            (setq new_val (itoa (1+ (atoi old_val))))
            (setq att_edata (subst (cons 1 new_val) (assoc 1 att_edata) att_edata))
            (entmod att_edata)
            (entupd ent)
            (setq count (1+ count))
          )
        )
        (setq i (1+ i))
      )
      (princ (strcat "\nSuccessfully incremented attributes for " (itoa count) " of " (itoa (sslength sel)) " selected blocks."))
    )
    (princ "\nNo block references with attributes were selected.")
  )
  (setvar "CMDECHO" old_cmdecho)
  (princ)
)

;; =============================================================================
;; == COMMAND D1: Decrement by 1
;; =============================================================================
(defun c:D1 (/ *error* old_cmdecho sel i count ent att_ent att_edata old_val new_val found_att)
  (defun *error* (msg)
    (if old_cmdecho (setvar "CMDECHO" old_cmdecho))
    (if (not (member msg '("Function cancelled" "quit / exit abort"))) (princ (strcat "\nError: " msg)))
    (princ)
  )
  (setq old_cmdecho (getvar "CMDECHO"))
  (setvar "CMDECHO" 0)
  (princ "\nSelect block(s) with attributes to decrement by 1: ")
  (setq sel (ssget '((0 . "INSERT") (66 . 1))))
  (if sel
    (progn
      (setq i 0 count 0)
      (repeat (sslength sel)
        (setq ent (ssname sel i) att_ent (entnext ent) found_att nil)
        (while (and att_ent (not found_att))
          (setq att_edata (entget att_ent))
          (if (= "ATTRIB" (cdr (assoc 0 att_edata)))
            (progn
              (setq old_val (cdr (assoc 1 att_edata)))
              (if (and old_val (numberp (read old_val))) (setq found_att T))
            )
          )
          (if (not found_att) (setq att_ent (entnext att_ent)))
        )
        (if found_att
          (progn
            ;; The only change is 1- instead of 1+
            (setq new_val (itoa (1- (atoi old_val))))
            (setq att_edata (subst (cons 1 new_val) (assoc 1 att_edata) att_edata))
            (entmod att_edata)
            (entupd ent)
            (setq count (1+ count))
          )
        )
        (setq i (1+ i))
      )
      (princ (strcat "\nSuccessfully decremented attributes for " (itoa count) " of " (itoa (sslength sel)) " selected blocks."))
    )
    (princ "\nNo block references with attributes were selected.")
  )
  (setvar "CMDECHO" old_cmdecho)
  (princ)
)

;; =============================================================================
;; == COMMAND C2: Increment by User Value
;; =============================================================================
(defun c:C2 (/ *error* old_cmdecho sel i count ent att_ent att_edata old_val new_val found_att inc_val)
  (defun *error* (msg)
    (if old_cmdecho (setvar "CMDECHO" old_cmdecho))
    (if (not (member msg '("Function cancelled" "quit / exit abort"))) (princ (strcat "\nError: " msg)))
    (princ)
  )
  (setq old_cmdecho (getvar "CMDECHO"))
  (setvar "CMDECHO" 0)
  (princ "\nSelect block(s) with attributes to increment: ")
  (setq sel (ssget '((0 . "INSERT") (66 . 1))))
  (if sel
    (progn
      ;; Prompt user for the increment value
      (setq inc_val (getint "\nEnter value to increment by: "))
      (if inc_val
        (progn
          (setq i 0 count 0)
          (repeat (sslength sel)
            (setq ent (ssname sel i) att_ent (entnext ent) found_att nil)
            (while (and att_ent (not found_att))
              (setq att_edata (entget att_ent))
              (if (= "ATTRIB" (cdr (assoc 0 att_edata)))
                (progn
                  (setq old_val (cdr (assoc 1 att_edata)))
                  (if (and old_val (numberp (read old_val))) (setq found_att T))
                )
              )
              (if (not found_att) (setq att_ent (entnext att_ent)))
            )
            (if found_att
              (progn
                ;; Use the user-provided value for incrementing
                (setq new_val (itoa (+ (atoi old_val) inc_val)))
                (setq att_edata (subst (cons 1 new_val) (assoc 1 att_edata) att_edata))
                (entmod att_edata)
                (entupd ent)
                (setq count (1+ count))
              )
            )
            (setq i (1+ i))
          )
          (princ (strcat "\nSuccessfully incremented attributes by " (itoa inc_val) " for " (itoa count) " of " (itoa (sslength sel)) " selected blocks."))
        )
        (princ "\nOperation cancelled. No value entered.")
      )
    )
    (princ "\nNo block references with attributes were selected.")
  )
  (setvar "CMDECHO" old_cmdecho)
  (princ)
)

;; =============================================================================
;; == COMMAND D2: Decrement by User Value
;; =============================================================================
(defun c:D2 (/ *error* old_cmdecho sel i count ent att_ent att_edata old_val new_val found_att dec_val)
  (defun *error* (msg)
    (if old_cmdecho (setvar "CMDECHO" old_cmdecho))
    (if (not (member msg '("Function cancelled" "quit / exit abort"))) (princ (strcat "\nError: " msg)))
    (princ)
  )
  (setq old_cmdecho (getvar "CMDECHO"))
  (setvar "CMDECHO" 0)
  (princ "\nSelect block(s) with attributes to decrement: ")
  (setq sel (ssget '((0 . "INSERT") (66 . 1))))
  (if sel
    (progn
      ;; Prompt user for the decrement value
      (setq dec_val (getint "\nEnter value to decrement by: "))
      (if dec_val
        (progn
          (setq i 0 count 0)
          (repeat (sslength sel)
            (setq ent (ssname sel i) att_ent (entnext ent) found_att nil)
            (while (and att_ent (not found_att))
              (setq att_edata (entget att_ent))
              (if (= "ATTRIB" (cdr (assoc 0 att_edata)))
                (progn
                  (setq old_val (cdr (assoc 1 att_edata)))
                  (if (and old_val (numberp (read old_val))) (setq found_att T))
                )
              )
              (if (not found_att) (setq att_ent (entnext att_ent)))
            )
            (if found_att
              (progn
                ;; Use the user-provided value for decrementing
                (setq new_val (itoa (- (atoi old_val) dec_val)))
                (setq att_edata (subst (cons 1 new_val) (assoc 1 att_edata) att_edata))
                (entmod att_edata)
                (entupd ent)
                (setq count (1+ count))
              )
            )
            (setq i (1+ i))
          )
          (princ (strcat "\nSuccessfully decremented attributes by " (itoa dec_val) " for " (itoa count) " of " (itoa (sslength sel)) " selected blocks."))
        )
        (princ "\nOperation cancelled. No value entered.")
      )
    )
    (princ "\nNo block references with attributes were selected.")
  )
  (setvar "CMDECHO" old_cmdecho)
  (princ)
)

;; =============================================================================
;; == COMMAND F1: Set attribute to a new user-entered value
;; =============================================================================
(defun c:F1 (/ *error* old_cmdecho sel i count ent att_ent att_edata new_val found_att)
  (defun *error* (msg)
    (if old_cmdecho (setvar "CMDECHO" old_cmdecho))
    (if (not (member msg '("Function cancelled" "quit / exit abort"))) (princ (strcat "\nError: " msg)))
    (princ)
  )
  (setq old_cmdecho (getvar "CMDECHO"))
  (setvar "CMDECHO" 0)
  (princ "\nSelect block(s) to set attribute value: ")
  (setq sel (ssget '((0 . "INSERT") (66 . 1))))
  (if sel
    (progn
      ;; Prompt user for the new string value. (getstring T) allows spaces.
      (setq new_val (getstring T "\nEnter the new value for the attribute: "))
      ;; Proceed only if the user entered a non-empty string
      (if (and new_val (/= new_val ""))
        (progn
          (setq i 0 count 0)
          (repeat (sslength sel)
            (setq ent (ssname sel i) att_ent (entnext ent) found_att nil)
            ;; Loop through the sub-entities of the block until the first attribute is found
            (while (and att_ent (not found_att))
              (setq att_edata (entget att_ent))
              (if (= "ATTRIB" (cdr (assoc 0 att_edata)))
                ;; We've found the first attribute, so set the flag to exit the while loop.
                (setq found_att T)
              )
              ;; If an attribute hasn't been found yet, move to the next sub-entity.
              (if (not found_att) (setq att_ent (entnext att_ent)))
            )
            ;; If an attribute was found for this block, modify it
            (if found_att
              (progn
                ;; Substitute the old value (dxf code 1) with the new user value
                (setq att_edata (subst (cons 1 new_val) (assoc 1 att_edata) att_edata))
                (entmod att_edata)
                (entupd ent)
                (setq count (1+ count))
              )
            )
            (setq i (1+ i))
          )
          (princ (strcat "\nSuccessfully set attribute to \"" new_val "\" for " (itoa count) " of " (itoa (sslength sel)) " selected blocks."))
        )
        (princ "\nOperation cancelled. No value entered.")
      )
    )
    (princ "\nNo block references with attributes were selected.")
  )
  (setvar "CMDECHO" old_cmdecho)
  (princ)
)


(princ "\nBlock Attribute Tools loaded. Commands available: C1, D1, C2, D2, F1.")