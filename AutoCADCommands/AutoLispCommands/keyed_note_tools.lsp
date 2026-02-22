; keyed_note_tools.lsp
; Commands:
;   KN       - Insert KEYED_NOTE in modelspace or active viewport and set keyed value

(vl-load-com)

(setq kn:*block-name* "KEYED_NOTE")
(setq kn:*layer-name* "keyed-notes")
(setq kn:*layer-color* 2) ; ACI 2 = yellow
(setq kn:*default-tag* "KEYNO")
(setq kn:*insert-scale* 1.0)
(setq kn:*modelspace-anno-scale* "1/4\" = 1'-0\"")
(setq kn:*canonical-text-style-name* "ARIALNARROW-1-8")
(setq kn:*canonical-text-style-font* "ARIALN.TTF")

; Canonical KEYED_NOTE definition extracted from:
; C:\Users\jakeh\Desktop\251126 BofA, 9591 Chapman Ave., Garden Grove, CA\Electrical\E02.00 Lighting.dwg
(setq kn:*canonical-base-point* '(-5.55112e-17 5.55112e-17 0.0))
(setq kn:*canonical-polyline-dxf*
  '((0 . "LWPOLYLINE")
    (100 . "AcDbEntity")
    (8 . "keyed-notes")
    (100 . "AcDbPolyline")
    (90 . 6)
    (70 . 129)
    (43 . 0.0)
    (38 . 0.0)
    (39 . 0.0)
    (10 -0.0625998 -0.108426)
    (40 . 0.0)
    (41 . 0.0)
    (42 . 0.0)
    (91 . 0)
    (10 0.0625998 -0.108426)
    (40 . 0.0)
    (41 . 0.0)
    (42 . 0.0)
    (91 . 0)
    (10 0.1252 -1.20737e-15)
    (40 . 0.0)
    (41 . 0.0)
    (42 . 0.0)
    (91 . 0)
    (10 0.0625998 0.108426)
    (40 . 0.0)
    (41 . 0.0)
    (42 . 0.0)
    (91 . 0)
    (10 -0.0625998 0.108426)
    (40 . 0.0)
    (41 . 0.0)
    (42 . 0.0)
    (91 . 0)
    (10 -0.1252 1.20737e-15)
    (40 . 0.0)
    (41 . 0.0)
    (42 . 0.0)
    (91 . 0)
    (210 0.0 0.0 1.0)))
(setq kn:*canonical-attdef-dxf*
  '((0 . "ATTDEF")
    (100 . "AcDbEntity")
    (8 . "keyed-notes")
    (100 . "AcDbText")
    (10 -0.107435 -0.046875 0.0)
    (40 . 0.09375)
    (1 . "XXX")
    (50 . 0.0)
    (41 . 1.0)
    (51 . 0.0)
    (7 . "ARIALNARROW-1-8")
    (71 . 0)
    (72 . 1)
    (11 0.0 3.53884e-15 0.0)
    (210 0.0 0.0 1.0)
    (100 . "AcDbAttributeDefinition")
    (280 . 0)
    (3 . "Enter No.")
    (2 . "2")
    (70 . 0)
    (73 . 0)
    (74 . 2)
    (280 . 1)))

(defun kn--item-or-nil (collection key / result)
  (setq result (vl-catch-all-apply 'vla-Item (list collection key)))
  (if (vl-catch-all-error-p result)
    nil
    result
  )
)

(defun kn--set-dxf (data code value / pair)
  (setq pair (assoc code data))
  (if pair
    (subst (cons code value) pair data)
    (append data (list (cons code value)))
  )
)

(defun kn--string-empty-p (value)
  (or (null value) (= "" (vl-string-trim " \t\r\n" value)))
)

(defun kn--clear-bit (value bit)
  (if (/= 0 (logand value bit))
    (- value bit)
    value
  )
)

(defun kn--join-types (types / out)
  (setq out "")
  (foreach typ types
    (setq out (strcat out typ ", "))
  )
  (if (> (strlen out) 2)
    (substr out 1 (- (strlen out) 2))
    out
  )
)

(defun kn--make-selection-set (enames / ss)
  (setq ss (ssadd))
  (foreach e enames
    (if (and e (entget e))
      (setq ss (ssadd e ss))
    )
  )
  ss
)

(defun kn--ensure-layer (doc / layers layerObj)
  (setq layers (vla-get-Layers doc))
  (setq layerObj (kn--item-or-nil layers kn:*layer-name*))
  (if (null layerObj)
    (setq layerObj (vla-Add layers kn:*layer-name*))
  )
  (if (/= (vla-get-Color layerObj) kn:*layer-color*)
    (vla-put-Color layerObj kn:*layer-color*)
  )
  layerObj
)

(defun kn--ensure-text-style (doc / styles styleObj)
  (setq styles (vla-get-TextStyles doc))
  (setq styleObj (kn--item-or-nil styles kn:*canonical-text-style-name*))
  (if (null styleObj)
    (progn
      (setq styleObj (vla-Add styles kn:*canonical-text-style-name*))
      (if (vlax-property-available-p styleObj 'FontFile T)
        (vl-catch-all-apply 'vla-put-FontFile (list styleObj kn:*canonical-text-style-font*))
      )
      (if (vlax-property-available-p styleObj 'BigFontFile T)
        (vl-catch-all-apply 'vla-put-BigFontFile (list styleObj ""))
      )
      (if (vlax-property-available-p styleObj 'Width T)
        (vl-catch-all-apply 'vla-put-Width (list styleObj 1.0))
      )
      (if (vlax-property-available-p styleObj 'ObliqueAngle T)
        (vl-catch-all-apply 'vla-put-ObliqueAngle (list styleObj 0.0))
      )
      (if (vlax-property-available-p styleObj 'Height T)
        (vl-catch-all-apply 'vla-put-Height (list styleObj 0.0))
      )
    )
  )
  styleObj
)

(defun kn--set-entity-layer-bylayer (ename / ed)
  (setq ed (entget ename))
  (setq ed (kn--set-dxf ed 8 kn:*layer-name*))
  (setq ed (kn--set-dxf ed 62 256))
  (entmod ed)
  (entupd ename)
)

(defun kn--set-annotative-if-possible (obj / result)
  (if (vlax-property-available-p obj 'Annotative T)
    (progn
      (setq result (vl-catch-all-apply 'vla-put-Annotative (list obj 1)))
      (if (vl-catch-all-error-p result)
        (vl-catch-all-apply 'vla-put-Annotative (list obj :vlax-true))
      )
    )
  )
)

(defun kn--closed-lwpolyline-p (ename / ed flags)
  (setq ed (entget ename))
  (setq flags (cdr (assoc 70 ed)))
  (= 1 (logand 1 (if flags flags 0)))
)

(defun kn--normalize-attdef (ename / ed flags tag)
  (setq ed (entget ename))
  (setq flags (cdr (assoc 70 ed)))
  (if (null flags)
    (setq flags 0)
  )
  ; Clear invisible and constant bits so the key remains visible and editable.
  (setq flags (kn--clear-bit flags 1))
  (setq flags (kn--clear-bit flags 2))
  (setq ed (kn--set-dxf ed 70 flags))

  (setq tag (cdr (assoc 2 ed)))
  (if (kn--string-empty-p tag)
    (setq ed (kn--set-dxf ed 2 kn:*default-tag*))
  )
  (if (kn--string-empty-p (cdr (assoc 3 ed)))
    (setq ed (kn--set-dxf ed 3 "Keyed Note"))
  )

  (setq ed (kn--set-dxf ed 8 kn:*layer-name*))
  (setq ed (kn--set-dxf ed 62 256))

  (entmod ed)
  (entupd ename)

  (kn--set-annotative-if-possible (vlax-ename->vla-object ename))
)

(defun kn--pick-entity (message expectedTypes / selection ename entityType done)
  (setq done nil)
  (setq ename nil)
  (while (not done)
    (setq selection (entsel message))
    (cond
      ((null selection)
        (setq done T)
      )
      (T
        (setq ename (car selection))
        (setq entityType (cdr (assoc 0 (entget ename))))
        (if (member entityType expectedTypes)
          (setq done T)
          (progn
            (prompt
              (strcat
                "\nInvalid selection. Expected one of: "
                (kn--join-types expectedTypes)
                "."
              )
            )
            (setq ename nil)
          )
        )
      )
    )
  )
  ename
)

(defun kn--get-block-for-use (doc / blockObj)
  (setq blockObj (kn--item-or-nil (vla-get-Blocks doc) kn:*block-name*))
  (if (null blockObj)
    nil
    (progn
      (if
        (or
          (and
            (vlax-property-available-p blockObj 'IsLayout)
            (= :vlax-true (vla-get-IsLayout blockObj))
          )
          (and
            (vlax-property-available-p blockObj 'IsXRef)
            (= :vlax-true (vla-get-IsXRef blockObj))
          )
        )
        nil
        blockObj
      )
    )
  )
)

(defun kn--variant-or-sa-to-list (value)
  (cond
    ((= (type value) 'VARIANT) (vlax-safearray->list (vlax-variant-value value)))
    ((= (type value) 'SAFEARRAY) (vlax-safearray->list value))
    (T nil)
  )
)

(defun kn--object-bbox-center (obj / minPt maxPt callResult minList maxList)
  (if (null obj)
    nil
    (progn
      (setq callResult (vl-catch-all-apply 'vla-GetBoundingBox (list obj 'minPt 'maxPt)))
      (if (vl-catch-all-error-p callResult)
        nil
        (progn
          (setq minList (kn--variant-or-sa-to-list minPt))
          (setq maxList (kn--variant-or-sa-to-list maxPt))
          (if (and (= 3 (length minList)) (= 3 (length maxList)))
            (list
              (/ (+ (nth 0 minList) (nth 0 maxList)) 2.0)
              (/ (+ (nth 1 minList) (nth 1 maxList)) 2.0)
              (/ (+ (nth 2 minList) (nth 2 maxList)) 2.0)
            )
            nil
          )
        )
      )
    )
  )
)

(defun kn--polyline-center-point (polyEname / polyObj)
  (setq polyObj (vlax-ename->vla-object polyEname))
  (kn--object-bbox-center polyObj)
)

(defun kn--object-array-variant (objects / sa idx)
  (setq sa (vlax-make-safearray vlax-vbObject (cons 0 (1- (length objects)))))
  (setq idx 0)
  (foreach obj objects
    (vlax-safearray-put-element sa idx obj)
    (setq idx (1+ idx))
  )
  (vlax-make-variant sa)
)

(defun kn--block-layout-or-xref-p (blockObj)
  (or
    (and
      (vlax-property-available-p blockObj 'IsLayout)
      (= :vlax-true (vla-get-IsLayout blockObj))
    )
    (and
      (vlax-property-available-p blockObj 'IsXRef)
      (= :vlax-true (vla-get-IsXRef blockObj))
    )
  )
)

(defun kn--clear-block-definition (blockObj / toDelete)
  (setq toDelete '())
  (vlax-for obj blockObj
    (setq toDelete (cons obj toDelete))
  )
  (foreach obj toDelete
    (vl-catch-all-apply 'vla-Delete (list obj))
  )
)

(defun kn--prepare-target-block (doc basePoint / blocks existing blockObj originResult)
  (setq blocks (vla-get-Blocks doc))
  (setq existing (kn--item-or-nil blocks kn:*block-name*))

  (if existing
    (if (kn--block-layout-or-xref-p existing)
      (setq blockObj nil)
      (progn
        ; Keep references intact and redefine content in place.
        (setq originResult
          (if (vlax-property-available-p existing 'Origin T)
            (vl-catch-all-apply 'vla-put-Origin (list existing (vlax-3d-point basePoint)))
            nil
          )
        )
        (if (and originResult (vl-catch-all-error-p originResult))
          (progn
            (prompt
              (strcat
                "\nWarning: Unable to set block origin directly: "
                (vl-catch-all-error-message originResult)
              )
            )
            (setq blockObj nil)
          )
          (progn
            (kn--clear-block-definition existing)
            (setq blockObj existing)
          )
        )
      )
    )
    (setq blockObj (vla-Add blocks (vlax-3d-point basePoint) kn:*block-name*))
  )

  blockObj
)

(defun kn--build-block-from-source (doc polyEname attEname / basePoint blockObj sourceObjects copyResult)
  ; Anchor inserts at the hexagon center so click point matches symbol center.
  (setq basePoint kn:*canonical-base-point*)
  (if (null basePoint)
    (progn
      (setq basePoint (kn--polyline-center-point polyEname))
      (if (null basePoint)
        (setq basePoint (cdr (assoc 10 (entget attEname))))
      )
    )
  )
  (if (null basePoint)
    (setq basePoint '(0.0 0.0 0.0))
  )

  (setq blockObj (kn--prepare-target-block doc basePoint))
  (if (null blockObj)
    (progn
      (prompt (strcat "\nUnable to create or redefine block " kn:*block-name* "."))
      nil
    )
    (progn
      (setq sourceObjects
        (list
          (vlax-ename->vla-object polyEname)
          (vlax-ename->vla-object attEname)
        )
      )
      (setq copyResult
        (vl-catch-all-apply
          'vlax-invoke-method
          (list doc 'CopyObjects (kn--object-array-variant sourceObjects) blockObj)
        )
      )
      (if (vl-catch-all-error-p copyResult)
        (progn
          (prompt
            (strcat
              "\nUnable to copy source objects into "
              kn:*block-name*
              ": "
              (vl-catch-all-error-message copyResult)
            )
          )
          nil
        )
        (progn
          (kn--set-annotative-if-possible blockObj)
          (vl-cmdf "_.ATTSYNC" "_N" kn:*block-name*)
          T
        )
      )
    )
  )
)

(defun kn--delete-entity-if-exists (ename)
  (if (and ename (entget ename))
    (entdel ename)
  )
)

(defun kn--build-block-from-canonical (doc / polyData attData polyEname attEname ok)
  (kn--ensure-text-style doc)
  (setq polyData (kn--set-dxf kn:*canonical-polyline-dxf* 8 kn:*layer-name*))
  (setq attData (kn--set-dxf kn:*canonical-attdef-dxf* 8 kn:*layer-name*))
  (setq attData (kn--set-dxf attData 7 kn:*canonical-text-style-name*))

  (setq polyEname (entmakex polyData))
  (setq attEname (entmakex attData))

  (if (and polyEname attEname)
    (progn
      (kn--set-entity-layer-bylayer polyEname)
      (kn--normalize-attdef attEname)
      (setq ok (kn--build-block-from-source doc polyEname attEname))
    )
    (setq ok nil)
  )

  (kn--delete-entity-if-exists polyEname)
  (kn--delete-entity-if-exists attEname)

  ok
)

(defun kn--ensure-canonical-keyed-note (doc)
  (kn--build-block-from-canonical doc)
)

(defun kn--valid-key-p (value / trimmed len idx code ok)
  (setq trimmed (vl-string-trim " \t\r\n" value))
  (if (= "" trimmed)
    nil
    (progn
      (setq len (strlen trimmed))
      (setq idx 1)
      (setq ok T)
      (while (and ok (<= idx len))
        (setq code (ascii (substr trimmed idx 1)))
        (if (or (< code 48) (> code 57))
          (setq ok nil)
        )
        (setq idx (1+ idx))
      )
      (and ok (> (atoi trimmed) 0))
    )
  )
)

(defun kn--prompt-key-value (/ value done)
  (setq value nil)
  (setq done nil)
  (while (not done)
    (setq value (getstring T "\nEnter keyed note number (e.g., 1, 2, 3): "))
    (cond
      ((or (null value) (= "" value))
        (setq value nil)
        (setq done T)
      )
      ((kn--valid-key-p value)
        (setq value (vl-string-trim " \t\r\n" value))
        (setq done T)
      )
      (T
        (prompt "\nInvalid value. Enter a positive whole number such as 1, 2, or 3.")
      )
    )
  )
  value
)

(defun kn--in-modelspace-p ()
  (or (= 1 (getvar "TILEMODE")) (> (getvar "CVPORT") 1))
)

(defun kn--in-viewport-editing-p ()
  (and (= 0 (getvar "TILEMODE")) (> (getvar "CVPORT") 1))
)

(defun kn--approx-eq (a b tol)
  (< (abs (- a b)) tol)
)

(defun kn--anno-scale-name-from-denom (denom / scaleMap name item)
  (setq scaleMap
    (list
      (cons 1 "1'-0\" = 1'-0\"")
      (cons 2 "6\" = 1'-0\"")
      (cons 4 "3\" = 1'-0\"")
      (cons 8 "1 1/2\" = 1'-0\"")
      (cons 12 "1\" = 1'-0\"")
      (cons 16 "3/4\" = 1'-0\"")
      (cons 24 "1/2\" = 1'-0\"")
      (cons 32 "3/8\" = 1'-0\"")
      (cons 48 "1/4\" = 1'-0\"")
      (cons 64 "3/16\" = 1'-0\"")
      (cons 96 "1/8\" = 1'-0\"")
      (cons 128 "3/32\" = 1'-0\"")
      (cons 192 "1/16\" = 1'-0\"")
      (cons 384 "1/32\" = 1'-0\"")
    )
  )
  (setq name nil)
  (foreach item scaleMap
    (if (= (car item) denom)
      (setq name (cdr item))
    )
  )
  (if name
    name
    (strcat "1:" (itoa denom))
  )
)

(defun kn--active-viewport-custom-scale (doc / vpObj value)
  (if (not (kn--in-viewport-editing-p))
    nil
    (progn
      (setq vpObj (vl-catch-all-apply 'vla-get-ActivePViewport (list doc)))
      (if (vl-catch-all-error-p vpObj)
        (progn
          (setq value (vl-catch-all-apply 'getvar (list "VPSCALE")))
          (if (vl-catch-all-error-p value) nil value)
        )
        (progn
          (setq value (vl-catch-all-apply 'vla-get-CustomScale (list vpObj)))
          (if (vl-catch-all-error-p value)
            (progn
              (setq value (vl-catch-all-apply 'getvar (list "VPSCALE")))
              (if (vl-catch-all-error-p value) nil value)
            )
            value
          )
        )
      )
    )
  )
)

(defun kn--target-anno-scale-for-viewport (doc / customScale denom rounded)
  (setq customScale (kn--active-viewport-custom-scale doc))
  (if (or (null customScale) (<= customScale 0.0))
    nil
    (progn
      (setq denom (/ 1.0 customScale))
      (setq rounded (fix (+ denom 0.5)))
      (if (kn--approx-eq denom rounded 1e-4)
        (kn--anno-scale-name-from-denom rounded)
        nil
      )
    )
  )
)

(defun kn--try-set-cannoscale (scaleName / result)
  (if (or (null scaleName) (= "" scaleName))
    nil
    (progn
      (setq result (vl-catch-all-apply 'setvar (list "CANNOSCALE" scaleName)))
      (not (vl-catch-all-error-p result))
    )
  )
)

(defun kn--first-new-insert-after (beforeEname / e ed found)
  (setq e (if beforeEname (entnext beforeEname) (entnext)))
  (while (and e (null found))
    (setq ed (entget e))
    (if (= "INSERT" (cdr (assoc 0 ed)))
      (setq found e)
    )
    (if (null found)
      (setq e (entnext e))
    )
  )
  found
)

(defun kn--set-insert-attributes (insertEname keyValue / e ed updated)
  (setq updated nil)
  (setq e (entnext insertEname))
  (while e
    (setq ed (entget e))
    (cond
      ((= "ATTRIB" (cdr (assoc 0 ed)))
        (setq ed (kn--set-dxf ed 1 keyValue))
        (entmod ed)
        (setq updated T)
      )
      ((= "SEQEND" (cdr (assoc 0 ed)))
        (setq e nil)
      )
    )
    (if e (setq e (entnext e)))
  )
  (if updated (entupd insertEname))
  updated
)

(defun kn--insert-keyed-note (doc insertPoint keyValue / oldAttReq oldAttDia beforeEnt result insertEname insObj)
  (setq oldAttReq (getvar "ATTREQ"))
  (setq oldAttDia (getvar "ATTDIA"))
  (setq beforeEnt (entlast))
  (setvar "ATTREQ" 0)
  (setvar "ATTDIA" 0)

  (setq result
    (vl-catch-all-apply
      'vl-cmdf
      (list "_.-INSERT" kn:*block-name* insertPoint kn:*insert-scale* kn:*insert-scale* 0.0)
    )
  )

  (setvar "ATTREQ" oldAttReq)
  (setvar "ATTDIA" oldAttDia)

  (if (vl-catch-all-error-p result)
    nil
    (progn
      (setq insertEname (kn--first-new-insert-after beforeEnt))
      (if insertEname
        (progn
          (kn--set-entity-layer-bylayer insertEname)
          (setq insObj (vlax-ename->vla-object insertEname))
          (if (vlax-property-available-p insObj 'XScaleFactor T)
            (vl-catch-all-apply 'vla-put-XScaleFactor (list insObj kn:*insert-scale*))
          )
          (if (vlax-property-available-p insObj 'YScaleFactor T)
            (vl-catch-all-apply 'vla-put-YScaleFactor (list insObj kn:*insert-scale*))
          )
          (if (vlax-property-available-p insObj 'ZScaleFactor T)
            (vl-catch-all-apply 'vla-put-ZScaleFactor (list insObj kn:*insert-scale*))
          )
          (kn--set-annotative-if-possible insObj)
          (if (not (kn--set-insert-attributes insertEname keyValue))
            (prompt "\nWarning: Inserted block has no editable attributes to update.")
          )
          insertEname
        )
        nil
      )
    )
  )
)

(defun c:KN (/ *error* doc oldCmdecho oldLayer targetAnnoScale keyValue insertPoint blockRef)
  (vl-load-com)
  (setq doc (vla-get-ActiveDocument (vlax-get-acad-object)))
  (setq oldCmdecho (getvar "CMDECHO"))
  (setq oldLayer (getvar "CLAYER"))
  (setvar "CMDECHO" 0)

  (defun *error* (msg)
    (if oldLayer (setvar "CLAYER" oldLayer))
    (if oldCmdecho (setvar "CMDECHO" oldCmdecho))
    (if
      (and msg (not (wcmatch (strcase msg) "*BREAK,*CANCEL*,*EXIT*")))
      (prompt (strcat "\nKN error: " msg))
    )
    (princ)
  )

  (cond
    ((not (kn--in-modelspace-p))
      (prompt "\nKN works in modelspace or inside an active layout viewport. Activate one and run again.")
    )
    (T
      (kn--ensure-layer doc)
      (if (kn--ensure-canonical-keyed-note doc)
        (progn
          (if (kn--in-viewport-editing-p)
            (progn
              (setq targetAnnoScale (kn--target-anno-scale-for-viewport doc))
              (if targetAnnoScale
                (if (kn--try-set-cannoscale targetAnnoScale)
                  (prompt (strcat "\nUsing viewport annotative scale: " targetAnnoScale))
                  (prompt
                    (strcat
                      "\nWarning: Could not set CANNOSCALE to "
                      targetAnnoScale
                      ". Using current annotation scale."
                    )
                  )
                )
                (prompt "\nWarning: Could not resolve active viewport scale to an annotation scale. Using current annotation scale.")
              )
            )
            (if (kn--try-set-cannoscale kn:*modelspace-anno-scale*)
              (prompt (strcat "\nUsing modelspace annotative scale: " kn:*modelspace-anno-scale*))
              (prompt
                (strcat
                  "\nWarning: Could not set CANNOSCALE to "
                  kn:*modelspace-anno-scale*
                  ". Using current annotation scale."
                )
              )
            )
          )
          (setq keyValue (kn--prompt-key-value))
          (if (null keyValue)
            (prompt "\nKN canceled.")
            (progn
              (setvar "CLAYER" kn:*layer-name*)
              (setq insertPoint (getpoint "\nSpecify keyed note insertion point: "))
              (if (null insertPoint)
                (prompt "\nKN canceled.")
                (progn
                  (setq blockRef (kn--insert-keyed-note doc insertPoint keyValue))
                  (if blockRef
                    (prompt
                      (strcat
                        "\nInserted "
                        kn:*block-name*
                        " with value "
                        keyValue
                        " on layer "
                        kn:*layer-name*
                        "."
                      )
                    )
                    (prompt "\nKN failed.")
                  )
                )
              )
            )
          )
        )
        (prompt (strcat "\nFailed to prepare canonical " kn:*block-name* " block definition."))
      )
    )
  )

  (if oldLayer (setvar "CLAYER" oldLayer))
  (setvar "CMDECHO" oldCmdecho)
  (princ)
)

(prompt "\nLoaded keyed_note_tools.lsp. Commands: KN.")
(princ)
