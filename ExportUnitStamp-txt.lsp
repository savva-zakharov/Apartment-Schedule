(defun c:UNITOUT (/ doc ms blk outname file atts attTags)
  (vl-load-com)

  ;; Active document + modelspace
  (setq doc (vla-get-ActiveDocument (vlax-get-acad-object)))
  (setq ms  (vla-get-ModelSpace doc))

  ;; Output file = DWG name + .txt
  (setq outname
        (strcat
          (getvar "DWGPREFIX")
          (vl-filename-base (getvar "DWGNAME"))
          ".txt"
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

  ;; Cache attribute tags for header
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
                       (list "HANDLE\tBLOCKNAME\t")
                       (mapcar
                         '(lambda (tag)
                            (strcat tag "\t")
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
                     (strcat "'" (vla-get-Handle blk) "\t")
                     (strcat (vla-get-EffectiveName blk) "\t")
                   )
                   (mapcar
                     '(lambda (attObj)
                        (strcat (vla-get-TextString attObj) "\t")
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
