;; LISP script for a single REV command.
;; FINAL VERSION - Spacebar confirms input (acts like Enter).
;; Uses AutoCAD's current default Revcloud Arc Lengths.

(defun create-revcloud-and-tag-final (num)
  ;; --- Setup & Layer Creation ---
  (setq rev_layer_name (strcat "rev-" num))
  (setq delta_layer_name (strcat "delta-" num))
  
  ;; Create layers if they don't exist
  (if (not (tblsearch "LAYER" rev_layer_name))
    (command "_.-LAYER" "_NEW" rev_layer_name "_COLOR" "6" rev_layer_name "")
  )
  (if (not (tblsearch "LAYER" delta_layer_name))
    (command "_.-LAYER" "_NEW" delta_layer_name "_COLOR" "6" delta_layer_name "")
  )

  ;; --- Part 1: Create the Revision Cloud ---
  ;; Uses current AutoCAD Arc Length settings.
  (initcommandversion 2)
  (command-s "_.REVCLOUD" "_R" PAUSE PAUSE)
  
  ;; Move the cloud to the correct layer
  (setq last_ent (entlast))
  (setq ent_data (entget last_ent))
  (setq new_data (subst (cons 8 rev_layer_name) (assoc 8 ent_data) ent_data))
  (entmod new_data)

  ;; --- Part 2: Interactively Place the Revision Delta Tag with Ghosting ---
  (prompt "\nSelect location for revision tag: ")
  
  ;; -- Define Triangle Geometry Constants --
  (setq side_length 0.296) 
  (setq h (* side_length (/ (sqrt 3.0) 2.0)))
  (setq R (* h (/ 2.0 3.0)))

  (setq tag_center nil) ; Initialize the final point
  (setq done nil)       ; Loop control flag

  (while (not done)
    (setq gr (grread T 15 0)) ; Read mouse/keyboard input
    (setq event_type (car gr))
    (setq current_point (cadr gr))

    (cond
      ;; If mouse is moved (event 5)
      ((= event_type 5)
        (redraw) ; Redraw to erase the previous ghost
        ; Calculate vertices relative to the current cursor position
        (setq p1 (polar current_point (* 0.5 pi) R))
        (setq p2 (polar current_point (* (/ 7.0 6.0) pi) R))
        (setq p3 (polar current_point (* (/ 11.0 6.0) pi) R))
        ; Draw the three sides of the triangle as temporary graphics
        (grdraw p1 p2 6 -1)
        (grdraw p2 p3 6 -1)
        (grdraw p3 p1 6 -1)
      )
      ;; If user clicks left mouse button (event 3)
      ((= event_type 3)
        (setq tag_center current_point) ; Lock in the final position
        (setq done T) ; Set flag to exit the loop
      )
      ;; If user hits Escape or right-clicks (event 2 or 25)
      ((or (= event_type 2) (= event_type 25))
        (setq done T) ; Set flag to exit the loop, tag_center remains nil
      )
    )
  )

  (redraw) ; Final redraw to clear the last ghost image

  ;; --- Create Permanent Objects if a point was selected ---
  (if tag_center
    (progn
      ;; Recalculate final vertices based on the chosen center point
      (setq p1 (polar tag_center (* 0.5 pi) R))
      (setq p2 (polar tag_center (* (/ 7.0 6.0) pi) R))
      (setq p3 (polar tag_center (* (/ 11.0 6.0) pi) R))

      ;; Draw the Triangle Polyline
      (entmake
        (list (cons 0 "LWPOLYLINE") (cons 100 "AcDbEntity") (cons 8 delta_layer_name)
              (cons 100 "AcDbPolyline") (cons 90 3) (cons 70 1)
              (cons 10 p1) (cons 10 p2) (cons 10 p3)
        )
      )
      
      ;; Create Text Style if missing
      (if (not (tblsearch "STYLE" "Arial"))
        (command "_.-STYLE" "Arial" "arial.ttf" "0.0" "" "" "" "")
      )
      
      ;; Draw the MText
      (entmake
        (list (cons 0 "MTEXT") (cons 100 "AcDbEntity") (cons 8 delta_layer_name)
              (cons 100 "AcDbMText") (cons 10 tag_center) (cons 40 0.1185)
              (cons 71 5) (cons 7 "Arial") (cons 1 num) (cons 41 0.0)
        )
      )
    )
  )
  (princ)
)

;; --- The Unified REV Command ---
(defun c:REV (/ user_val prompt_msg)
  ;; Initialize global variable for remembering last used revision if not set
  (if (null *last_rev_val*)
      (setq *last_rev_val* "1") ; Default to 1 if first time running
  )

  ;; Create the prompt string showing the default value
  (setq prompt_msg (strcat "\nEnter revision value <" *last_rev_val* ">: "))
  
  ;; Ask user for input. 
  ;; REMOVED the "T" flag. This means Spacebar now terminates input (same as Enter).
  (setq user_val (getstring prompt_msg))

  ;; Logic: If user hits Enter/Space (empty string), use the previous value
  (if (= user_val "")
      (setq user_val *last_rev_val*)
      (setq *last_rev_val* user_val) ; Update the memory with the new input
  )

  ;; Run the core function
  (create-revcloud-and-tag-final user_val)
  (princ)
)

(princ "\nCommand 'REV' loaded. Spacebar now confirms input.")
(princ)