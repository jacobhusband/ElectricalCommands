;; ElectricalCommands AutoLISP Loader
;; Loads all bundled LSP utilities safely.

(vl-load-com)
(setq ec-loader-file "ElectricalCommands.AutoLispCommands.lsp")

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

;; Discover utility files in the bundle folder and exclude the loader itself.
(defun ec-discover-lisp-files ( / files)
  (if (= (type ec-base-dir) 'STR)
    (progn
      (setq files (vl-directory-files ec-base-dir "*.lsp" 1))
      (if files
        (setq files
          (vl-remove-if
            '(lambda (x) (= (strcase x) (strcase ec-loader-file)))
            files
          )
        )
      )
      (if files
        (acad_strlsort files)
        nil
      )
    )
    nil
  )
)

(setq ec-autolisp-files (ec-discover-lisp-files))

(if ec-autolisp-files
  (progn
    (setq ec-total-count (length ec-autolisp-files))
    (setq ec-loaded-count 0)
    (princ (strcat "\n[ElectricalCommands] Discovered " (itoa ec-total-count) " utility .lsp file(s)."))

    (foreach f ec-autolisp-files
      (if (ec-load-lisp f)
        (setq ec-loaded-count (+ ec-loaded-count 1))
      )
    )

    (princ
      (strcat
        "\nElectricalCommands AutoLISP commands loaded ("
        (itoa ec-loaded-count)
        "/"
        (itoa ec-total-count)
        ")."
      )
    )
  )
  (princ "\n[ElectricalCommands] Warning: No utility .lsp files discovered to load.")
)

(princ)
