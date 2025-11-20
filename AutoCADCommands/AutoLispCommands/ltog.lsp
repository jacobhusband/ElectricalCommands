(defun c:LTOG ( / searchString ucSearchString layerData layerName ucLayerName layersToProcess layerString action firstLayerData)
  ;; Prompt the user for a string to search for
  (setq searchString (getstring T "\nEnter text string for XREF layers to toggle freeze: "))

  (if (and searchString (> (strlen searchString) 0))
    (progn
      (setq ucSearchString (strcat "*" (strcase searchString) "*"))
      (setq layersToProcess nil) ; Initialize list for matching layer names

      ;; 1. Collect all matching layers
      (setq layerData (tblnext "LAYER" T))
      (while layerData
        (setq layerName (cdr (assoc 2 layerData)))
        (setq ucLayerName (strcase layerName))
        (if (and (wcmatch layerName "*|*") (wcmatch ucLayerName ucSearchString))
          (setq layersToProcess (cons layerName layersToProcess))
        )
        (setq layerData (tblnext "LAYER"))
      )

      ;; 2. Determine action based on the first layer found
      (if layersToProcess
        (progn
          ;; Get the properties of the first layer in our list
          (setq firstLayerData (tblsearch "LAYER" (car layersToProcess)))
          ;; Check its frozen status (70 code, bit 1)
          (if (= 1 (logand 1 (cdr (assoc 70 firstLayerData))))
            (setq action "_THAW")  ; If it's frozen, we will thaw
            (setq action "_FREEZE") ; If it's thawed, we will freeze
          )
          
          ;; 3. Execute the action
          (setq layerString (apply 'strcat (mapcar '(lambda (x) (strcat x ",")) layersToProcess)))
          (setq layerString (substr layerString 1 (1- (strlen layerString)))) ; Remove last comma
          (command "_.-LAYER" action layerString "")
          
          (princ (strcat "\n" (if (= action "_THAW") "Thawed " "Froze ") (itoa (length layersToProcess)) " matching XREF layer(s)."))
        )
        (princ (strcat "\nNo XREF layers found containing \"" searchString "\"."))
      )
    )
    (princ "\nNo search string entered. Command cancelled.")
  )
  (princ)
)