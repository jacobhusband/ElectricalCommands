(defun c:SumLengths (/ ss total-length i ent ent-type vla-obj)
  (princ "\nSelect polylines and lines to sum their lengths.")
  ;; Updated ssget to include LWPOLYLINE
  (setq ss (ssget '((0 . "LWPOLYLINE,POLYLINE,LINE"))))
  (if ss
    (progn
      (setq total-length 0.0)
      (setq i 0)
      (repeat (sslength ss)
        (setq ent (ssname ss i))
        (setq ent-type (cdr (assoc 0 (entget ent))))
        ;; Check for all three types
        (if (or (= ent-type "LWPOLYLINE") (= ent-type "POLYLINE") (= ent-type "LINE"))
          (progn
            (setq vla-obj (vlax-ename->vla-object ent))
            (setq total-length (+ total-length (vlax-get-property vla-obj 'Length)))
          )
        )
        (setq i (1+ i))
      )
      (princ (strcat "\nTotal length of selected objects: " (rtos total-length 2 4)))
    )
    (princ "\nNo polylines or lines were selected.")
  )
  (princ)
)

;; A shorter alias for the command
(defun c:SL () (c:SumLengths))