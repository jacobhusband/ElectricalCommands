(defun str-similarity (s1 s2 / l1 l2 c n i j)
  ;; crude similarity: count of matching chars in order, normalized
  (setq s1 (strcase s1) s2 (strcase s2))
  (setq l1 (strlen s1) l2 (strlen s2) c 0 n 0 i 1 j 1)
  (while (and (<= i l1) (<= j l2))
    (if (= (substr s1 i 1) (substr s2 j 1))
      (progn (setq c (1+ c)) (setq j (1+ j)) (setq i (1+ i)))
      (setq j (1+ j))
    )
    (if (> j l2) (setq i (1+ i) j 1))
  )
  (if (> (max l1 l2) 0)
    (/ (float c) (max l1 l2))
    0.0
  )
)

(defun best-layer (pattern / tbl lay bestLayer bestScore score)
  ;; search for closest layer to pattern
  ;; Initialize with current layer as a fallback
  (setq bestLayer (getvar "CLAYER")
        bestScore 0.0
        tbl (tblnext "LAYER" T)
  )
  (while tbl
    (setq lay (cdr (assoc 2 tbl)))
    (setq score (str-similarity pattern lay))
    (if (> score bestScore)
      (progn
        (setq bestScore score)
        (setq bestLayer lay)
      )
    )
    (setq tbl (tblnext "LAYER"))
  )
  bestLayer
)

(defun c:CHLAY ( / ss pattern targetLayer )
  ;; First, check for a pre-existing selection set
  (setq ss (ssget "_I"))
  
  (setq pattern (getstring T "\nEnter layer search string: "))
  (setq targetLayer (best-layer pattern))

  (if (null targetLayer)
    (princ "\nNo matching layer found.")
    (progn
      (if ss
        (progn
          ;; Use the "Previous" selection set in the CHPROP command
          (command "_.CHPROP" "_P" "" "_LA" targetLayer "")
          (princ (strcat "\nObjects moved to layer: " targetLayer))
        )
        (progn
          ;; No selection, so change the current layer
          (setvar "CLAYER" targetLayer)
          (princ (strcat "\nCurrent layer set to: " targetLayer))
        )
      )
    )
  )
  (princ)
)
