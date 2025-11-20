(defun c:MTOG ( / searchList layersToProcess layerData layerName ucLayerName layerString action firstLayerData)
  (setq searchList '("christian" "CD stamp"))
  (setq layersToProcess nil)

  ;; 1) Collect all layers whose names contain any search term (case-insensitive)
  (setq layerData (tblnext "LAYER" T))
  (while layerData
    (setq layerName   (cdr (assoc 2 layerData)))
    (setq ucLayerName (strcase layerName))
    (if (vl-some '(lambda (s) (wcmatch ucLayerName (strcat "*" (strcase s) "*"))) searchList)
      (setq layersToProcess (cons layerName layersToProcess))
    )
    (setq layerData (tblnext "LAYER"))
  )

  ;; 2) Determine action from first match, then apply to all
  (if layersToProcess
    (progn
      (setq firstLayerData (tblsearch "LAYER" (car layersToProcess)))
      (if (= 1 (logand 1 (cdr (assoc 70 firstLayerData))))
        (setq action "_THAW")
        (setq action "_FREEZE")
      )

      ;; Build comma-separated list for -LAYER
      (setq layerString
            (vl-string-right-trim ","
              (apply 'strcat (mapcar '(lambda (x) (strcat x ",")) (reverse layersToProcess)))))

      (command "_.-LAYER" action layerString "")

      (princ (strcat
              "\n" (if (= action "_THAW") "Thawed " "Froze ")
              (itoa (length layersToProcess)) " layer(s) matching criteria."))
    )
    (princ "\nNo matching layers found for \"christian\" or \"CD stamp\".")
  )
  (princ)
)
