(defun c:ExportUnitStamp (/ doc ms blk outname file atts attTags
                            sep quote fmt)
  (vl-load-com)

  ;; Ask for format (default TXT)
	;; Prompt for format
	(setq fmt (getstring "\nExport format [TXT/CSV] <TXT>: "))
	(if (or (null fmt) (= fmt "")) (setq fmt "TXT"))  ; default if Enter pressed
	(setq fmt (strcase fmt))


  ;; Set separator + quoting
  (cond
    ((= fmt "CSV")
     (setq sep ",")
     (setq quote T))
    (T
     (setq sep "\t")
     (setq quote nil))
  )

  ;; Active document + modelspace
  (setq doc (vla-get-ActiveDocument (vlax-get-acad-object)))
  (setq ms  (vla-get-ModelSpace doc))

  ;; Output file name
  (setq outname
        (strcat
          (getvar "DWGPREFIX")
          (vl-filename-base (getvar "DWGNAME"))
          (if (= fmt "CSV") ".csv" ".txt")
        )
  )

  ;; Open file (overwrite)
  (setq file (open outname "w"))
  (if (not file)
    (progn
      (princ "\nUnable to open output file.")
      (exit)
    )
  )

  ;; Cache attribute tags
  (setq attTags nil)

  ;; Iterate ModelSpace
  (vlax-for blk ms
    (if (and
          (= (vla-get-ObjectName blk) "AcDbBlockReference")
          (= (strcase (vla-get-EffectiveName blk)) "00-UNIT-STAMP")
          (vla-get-HasAttributes blk)
        )
      (progn
        (setq atts (vlax-invoke blk 'GetAttributes))

        ;; Write header once
        (if (null attTags)
          (progn
            (setq attTags
                  (mapcar
                    '(lambda (tagObj)
                       (vla-get-TagString tagObj)
                     )
                    atts
                  )
            )
            (write-line
              (apply 'strcat
                     (append
                       (list
                         (strcat
                           "HANDLE" sep
                           "BLOCKNAME" sep
                         )
                       )
                       (mapcar
                         '(lambda (tag)
                            (strcat
                              (if quote "\"" "")
                              tag
                              (if quote "\"" "")
                              sep
                            )
                          )
                         attTags
                       )
                     )
              )
              file
            )
          )
        )

        ;; Write data row
        (write-line
          (apply 'strcat
                 (append
                   (list
                     (strcat
                       (if quote "\"" "")
                       "'" (vla-get-Handle blk)
                       (if quote "\"" "")
                       sep
                     )
                     (strcat
                       (if quote "\"" "")
                       (vla-get-EffectiveName blk)
                       (if quote "\"" "")
                       sep
                     )
                   )
                   (mapcar
                     '(lambda (attObj)
                        (strcat
                          (if quote "\"" "")
                          (vla-get-TextString attObj)
                          (if quote "\"" "")
                          sep
                        )
                      )
                     atts
                   )
                 )
          )
          file
        )
      )
    )
  )

  (close file)

  (princ
    (strcat
      "\nAttributes exported to: "
      outname
    )
  )

  (princ)
)
