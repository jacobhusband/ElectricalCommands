;; LISP script for an enhanced REV command.
;; FINAL VERSION - Fixed "Rectangular" Error.
;; Breaks command inputs into separate steps to force Modern REVCLOUD behavior.

(defun c:REV (/ input loop start_pt mode ent_before ent_data new_data rev_layer_name delta_layer_name temp_val approx_arc min_arc max_arc last_ent env_val env_arc)
  ;; --- 1. Memory Setup ---
  (if (null *last_rev_val*)
    (progn
      (setq env_val (getenv "REV_LAST_VAL"))
      (if (and env_val (/= env_val ""))
        (setq *last_rev_val* env_val)
        (setq *last_rev_val* "1")
      )
    )
  )

  (if (null *last_rev_arc*)
    (progn
      (setq env_arc (getenv "REV_ARC_APPROX"))
      (if (and env_arc (/= env_arc ""))
        (setq *last_rev_arc* (atof env_arc))
      )
    )
  )

  (if (and *last_rev_arc* (> *last_rev_arc* 0.0))
    (progn
      (setq min_arc (* *last_rev_arc* 0.8))
      (setq max_arc (* *last_rev_arc* 1.2))
      (command "_.REVCLOUD" "_A" min_arc max_arc)
      (command) ; Cancel immediately to save settings for this drawing
    )
  )
  
  (setq loop T)
  (setq mode "_R") ; Default mode
  (setq start_pt nil)

  ;; --- 2. The Main Prompt Loop ---
  (while loop
    (initget "Value Arc Rectangular Polygonal Freehand Object Style Modify")
    
    (setq input 
      (getpoint (strcat "\nCurrent Rev: " *last_rev_val* " - Specify start point or [Value/Arc/Rectangular/Polygonal/Freehand/Object/Style/Modify] <Rectangular>: "))
    )
    
    (cond
      ;; --- Option: Value ---
      ((= input "Value")
       (setq temp_val (getstring (strcat "\nEnter new revision value <" *last_rev_val* ">: ")))
       (if (/= temp_val "") 
           (progn
             (setq *last_rev_val* temp_val)
             (setenv "REV_LAST_VAL" *last_rev_val*)
           )
       )
      )

      ;; --- Option: Arc Length (Custom Approx Logic) ---
      ((= input "Arc")
       (setq approx_arc (getdist "\nSpecify approximate length of arc: "))
       (if approx_arc
         (progn
           (setq *last_rev_arc* approx_arc)
           (setenv "REV_ARC_APPROX" (rtos *last_rev_arc* 2 4))
           (setq min_arc (* approx_arc 0.8))
           (setq max_arc (* approx_arc 1.2))
           (command "_.REVCLOUD" "_A" min_arc max_arc) 
           (command) ; Cancel immediately to save settings
           (prompt (strcat "\nArc length set (Approx: " (rtos approx_arc 2 2) ")."))
         )
       )
      )

      ;; --- Drawing Modes & Style ---
      ((= input "Style")       (setq mode "_S") (setq loop nil))
      ((= input "Rectangular") (setq mode "_R") (setq loop nil))
      ((= input "Polygonal")   (setq mode "_P") (setq loop nil))
      ((= input "Freehand")    (setq mode "_F") (setq loop nil))
      ((= input "Object")      (setq mode "_O") (setq loop nil))
      ((= input "Modify")      (setq mode "_M") (setq loop nil))

      ;; --- User Clicked a Point ---
      ((listp input)
       (setq start_pt input)
       (setq mode "_R")
       (setq loop nil)
      )

      ;; --- User hit Enter ---
      ((null input)
       (setq mode "_R")
       (setq loop nil)
      )
    )
  )

  ;; --- 3. Execute Drawing & Layer Management ---
  (setq rev_layer_name (strcat "rev-" *last_rev_val*))
  (setq delta_layer_name (strcat "delta-" *last_rev_val*))

  (if (not (tblsearch "LAYER" rev_layer_name))
    (command "_.-LAYER" "_NEW" rev_layer_name "_COLOR" "6" rev_layer_name "")
  )
  (if (not (tblsearch "LAYER" delta_layer_name))
    (command "_.-LAYER" "_NEW" delta_layer_name "_COLOR" "6" delta_layer_name "")
  )

  (setq ent_before (entlast))

  ;; --- CRITICAL FIX: Split Command Calls ---
  ;; 1. Initialize Modern Version
  (initcommandversion 2)
  
  ;; 2. Start Command EMPTY. This forces AutoCAD to load the modern UI.
  (command "_.REVCLOUD")
  
  ;; 3. Feed options based on logic
  (cond
    ;; Case A: Rectangular with a pre-selected start point
    ((and (= mode "_R") start_pt)
      (command "_R")       ; Switch to Rectangular
      (command start_pt)   ; Feed the point we clicked earlier
    )
    
    ;; Case B: Just Rectangular (no point yet)
    ((= mode "_R")
      (command "_R")
    )

    ;; Case C: Any other mode (Polygonal, Object, etc)
    (T
      (command mode)
    )
  )

  ;; 4. Let user finish the command interactively
  (while (> (getvar "CMDACTIVE") 0) 
    (command PAUSE)
  )

  ;; --- 4. Tagging ---
  (setq last_ent (entlast))

  (if (not (equal ent_before last_ent))
    (progn
      (setq ent_data (entget last_ent))
      (setq new_data (subst (cons 8 rev_layer_name) (assoc 8 ent_data) ent_data))
      (entmod new_data)
      (create-rev-tag-only *last_rev_val* delta_layer_name)
    )
  )
  (princ)
)

;; --- Helper Function: Tag ---
(defun create-rev-tag-only (num layer_name)
  (prompt "\nSelect location for revision tag: ")
  (setq side_length 0.296) 
  (setq h (* side_length (/ (sqrt 3.0) 2.0)))
  (setq R (* h (/ 2.0 3.0)))
  (setq tag_center nil)
  (setq done nil)
  (while (not done)
    (setq gr (grread T 15 0))
    (setq event_type (car gr))
    (setq current_point (cadr gr))
    (cond
      ((= event_type 5)
        (redraw) 
        (setq p1 (polar current_point (* 0.5 pi) R))
        (setq p2 (polar current_point (* (/ 7.0 6.0) pi) R))
        (setq p3 (polar current_point (* (/ 11.0 6.0) pi) R))
        (grdraw p1 p2 6 -1) (grdraw p2 p3 6 -1) (grdraw p3 p1 6 -1)
      )
      ((= event_type 3) (setq tag_center current_point) (setq done T))
      ((or (= event_type 2) (= event_type 25)) (setq done T))
    )
  )
  (redraw) 
  (if tag_center
    (progn
      (setq p1 (polar tag_center (* 0.5 pi) R))
      (setq p2 (polar tag_center (* (/ 7.0 6.0) pi) R))
      (setq p3 (polar tag_center (* (/ 11.0 6.0) pi) R))
      (entmake (list (cons 0 "LWPOLYLINE") (cons 100 "AcDbEntity") (cons 8 layer_name) (cons 100 "AcDbPolyline") (cons 90 3) (cons 70 1) (cons 10 p1) (cons 10 p2) (cons 10 p3)))
      (if (not (tblsearch "STYLE" "Arial")) (command "_.-STYLE" "Arial" "arial.ttf" "0.0" "" "" "" ""))
      (entmake (list (cons 0 "MTEXT") (cons 100 "AcDbEntity") (cons 8 layer_name) (cons 100 "AcDbMText") (cons 10 tag_center) (cons 40 0.1185) (cons 71 5) (cons 7 "Arial") (cons 1 num) (cons 41 0.0)))
    )
  )
)
(princ "\nCommand 'REV' loaded.")
(princ)
