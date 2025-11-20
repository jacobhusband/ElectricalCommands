;; ElectricalCommands AutoLISP Loader
;; Loads all bundled LSP utilities safely.

(vl-load-com)

;; Function to safely find the bundle directory based on this file's location
(defun ec-get-bundle-dir ( / this-file full-path dir)
  ;; 1. Try to find this loader file itself
  (setq this-file "ElectricalCommands.AutoLispCommands.lsp")
  (setq full-path (findfile this-file))
  
  (cond
    ;; Case A: File found and is a string
    ((= (type full-path) 'STR)
      (setq dir (vl-filename-directory full-path))
      (princ (strcat "\n[ElectricalCommands] Bundle Root detected: " dir))
      dir ;; Return directory string
    )
    ;; Case B: Not found, return nil
    (t 
      (princ "\n[ElectricalCommands] Warning: Could not detect bundle root path.")
      nil
    )
  )
)

;; Initialize the base directory variable
(setq ec-base-dir (ec-get-bundle-dir))

;; Function to load a specific file
(defun ec-load-lisp (fname / full-path loaded)
  ;; 1. Try to find file using standard AutoCAD support paths
  (setq full-path (findfile fname))
  
  ;; 2. If not found, and we know the bundle dir, look there explicitly
  (if (and (null full-path) (= (type ec-base-dir) 'STR))
    (setq full-path (findfile (strcat ec-base-dir "\\" fname)))
  )

  ;; 3. Load it if found
  (if (and full-path (= (type full-path) 'STR))
    (progn
      ;; Use vl-catch-all-apply to prevent one bad file from crashing the whole loader
      (if (vl-catch-all-error-p (vl-catch-all-apply 'load (list full-path)))
        (princ (strcat "\n[ElectricalCommands] ERROR loading file: " fname))
        (setq loaded T)
      )
    )
    (princ (strcat "\n[ElectricalCommands] Missing file: " fname))
  )
  loaded
)

;; List of files to load
(setq ec-autolisp-files
      '("blockIncrementDecrement.lsp"
        "chlay.lsp"
        "cleanup.lsp"
        "etog.lsp"
        "ltog.lsp"
        "mtog.lsp"
        "revclouds.lsp"
        "sumlengths.lsp"
        "textcount.lsp"
        "tselect.lsp")
)

;; Load loop
(foreach f ec-autolisp-files
  (ec-load-lisp f)
)

(princ "\nElectricalCommands AutoLISP commands loaded.")
(princ)