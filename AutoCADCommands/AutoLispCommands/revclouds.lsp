;; LISP script for a two-part revision command.
;; FINAL VERSION - Includes a live "ghost" preview of the delta tag that follows the mouse.

(defun create-revcloud-and-tag-final (num)
  ;; --- Setup & Layer Creation ---
  (setq rev_layer_name (strcat "rev-" num))
  (setq delta_layer_name (strcat "delta-" num))
  (if (not (tblsearch "LAYER" rev_layer_name))
    (command "_.-LAYER" "_NEW" rev_layer_name "_COLOR" "6" rev_layer_name "")
  )
  (if (not (tblsearch "LAYER" delta_layer_name))
    (command "_.-LAYER" "_NEW" delta_layer_name "_COLOR" "6" delta_layer_name "")
  )

  ;; --- Part 1: Create the Revision Cloud ---
  (initcommandversion 2)
  (command-s "_.REVCLOUD" "_A" "3/16\"" "1/4\"" "_R" PAUSE PAUSE)
  (setq last_ent (entlast))
  (setq ent_data (entget last_ent))
  (setq new_data (subst (cons 8 rev_layer_name) (assoc 8 ent_data) ent_data))
  (entmod new_data)

  ;; --- Part 2: Interactively Place the Revision Delta Tag with Ghosting ---
  (prompt "\nSelect location for revision tag: ")
  
  ;; -- Define Triangle Geometry Constants --
  (setq side_length 0.296) ; <-- This value has been changed from 0.37
  (setq h (* side_length (/ (sqrt 3.0) 2.0)))
  (setq R (* h (/ 2.0 3.0)))

  (setq tag_center nil) ; Initialize the final point
  (setq done nil)      ; Loop control flag

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

      (entmake
        (list (cons 0 "LWPOLYLINE") (cons 100 "AcDbEntity") (cons 8 delta_layer_name)
              (cons 100 "AcDbPolyline") (cons 90 3) (cons 70 1)
              (cons 10 p1) (cons 10 p2) (cons 10 p3)
        )
      )
      (if (not (tblsearch "STYLE" "Arial"))
        (command "_.-STYLE" "Arial" "arial.ttf" "0.0" "" "" "" "")
      )
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

;; --- Define the Commands rev1 through rev9 ---
(defun c:rev1 () (create-revcloud-and-tag-final "1"))
(defun c:rev2 () (create-revcloud-and-tag-final "2"))
(defun c:rev3 () (create-revcloud-and-tag-final "3"))
(defun c:rev4 () (create-revcloud-and-tag-final "4"))
(defun c:rev5 () (create-revcloud-and-tag-final "5"))
(defun c:rev6 () (create-revcloud-and-tag-final "6"))
(defun c:rev7 () (create-revcloud-and-tag-final "7"))
(defun c:rev8 () (create-revcloud-and-tag-final "8"))
(defun c:rev9 () (create-revcloud-and-tag-final "9"))

;; --- Confirmation Message ---
(princ "\nRevision commands rev1 to rev9 (with live preview) loaded successfully.")
(princ)