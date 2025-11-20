(defun c:TEXTCOUNT (/ result counts total)
  (vl-load-com)
  (setq result (_tc-collect)) ; returns (list counts total) or nil
  (if result
    (progn
      (setq counts (car result)
            total  (cadr result))

      ;; List results in the command line
      (princ "\n--- Text Object Count ---")
      (foreach pair counts
        (princ (strcat "\n" (car pair) " : " (itoa (cdr pair)))))
      (princ (strcat "\n-------------------------"
                     "\nTotal unique text objects: " (itoa (length counts))
                     "\nTotal objects processed: " (itoa total)))

      ;; Write to file on desktop
      (if (_tc-write-file-to-desktop counts total)
        (princ "\n\nExported counts to TXT and opened in Notepad.")
        (princ "\n\nCould not write file to desktop.")
      )
    )
    (princ "\nNo TEXT/MTEXT/ATTRIB selected."))
  (princ)
)

(defun _tc-collect (/ ss n obj raw val counts pair)
  (if (setq ss (ssget '((0 . "TEXT,MTEXT,ATTRIB"))))
    (progn
      (setq counts '() n 0)
      (while (< n (sslength ss))
        (setq obj (vlax-ename->vla-object (ssname ss n)))
        (setq raw (vla-get-TextString obj))
        (if raw
          (progn
            (setq val (vl-string-subst " " "\\P" raw)) ; replace MTEXT line breaks
            (setq val (vl-string-trim " \t\r\n" val))  ; trim spaces
            (if (/= val "")
              (if (setq pair (assoc val counts))
                (setq counts (subst (cons val (1+ (cdr pair))) pair counts))
                (setq counts (cons (cons val 1) counts))
              )
            )
          )
        )
        (setq n (1+ n))
      )
      ;; Sort by count desc, then alphabetically
      (setq counts
        (vl-sort counts
          (function (lambda (a b)
            (if (/= (cdr a) (cdr b))
              (> (cdr a) (cdr b))
              (< (strcase (car a)) (strcase (car b)))
            )))))
      (list counts (sslength ss))
    )
  )
)

(defun _tc-write-file-to-desktop (counts total / desktop-path path f)
  ;; Get the desktop path from the Windows registry
  (setq desktop-path (vl-registry-read "HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\User Shell Folders" "Desktop"))
  
  (if (and desktop-path (setq path (strcat
                                     desktop-path
                                     "\\TextCount_"
                                     (menucmd "M=$(edtime,$(getvar,DATE),YYYY-MM-DD_HH-MM-SS)")
                                     ".txt")))
    (if (setq f (open path "w"))
      (progn
        (write-line (strcat "Total selected text objects: " (itoa total)) f)
        (write-line "" f)
        (foreach p counts
          (write-line (strcat (car p) " : " (itoa (cdr p))) f))
        (close f)
        (startapp "notepad.exe" path)
        (princ (strcat "\nWrote: " path))
        t
      )
      (progn (princ (strcat "\nERROR: Could not open file for writing at: " path)) nil)
    )
    (progn (princ "\nERROR: Could not determine desktop path.") nil)
  )
)

;; The _tc-make-table function is no longer called, but can be kept for future use if needed.
(defun _tc-make-table (counts insPt / acad doc space rows cols txtsize rowH colW1 colW2 tbl i)
  (vl-load-com)
  (setq acad (vlax-get-acad-object)
        doc  (vla-get-ActiveDocument acad)
        space (if (= (getvar "TILEMODE") 1)
                 (vla-get-ModelSpace doc)
                 (if (= (getvar "CVPORT") 1)
                   (vla-get-PaperSpace doc)
                   (vla-get-ModelSpace doc))))
  ;; Sizing
  (setq txtsize (getvar "TEXTSIZE"))
  (if (or (not txtsize) (<= txtsize 0.0)) (setq txtsize 2.5))
  (setq rowH  (* txtsize 1.6))
  (setq colW1 (* txtsize 16.0))
  (setq colW2 (* txtsize 6.0))

  (setq rows (1 + (length counts))) ; header + data
  (setq cols 2)

  (setq tbl (vla-AddTable space (vlax-3d-point insPt) rows cols rowH colW1))

  ;; Headers
  (vla-SetText tbl 0 0 "Text")
  (vla-SetText tbl 0 1 "Count")

  ;; Fill rows
  (setq i 0)
  (foreach p counts
    (vla-SetText tbl (1+ i) 0 (car p))
    (vla-SetText tbl (1+ i) 1 (itoa (cdr p)))
    (setq i (1+ i))
  )
  t
)