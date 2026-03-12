(defun c:UNITOUT (/ doc ss i ent obj outname file atts attTags
                            attPairs sep quote fmt val tag)

  (vl-load-com)

  ;; Prompt for format
  (setq fmt (getstring "\nExport format [TXT/CSV] <CSV>: "))
  (if (or (null fmt) (= fmt "")) (setq fmt "TXT"))
  (setq fmt (strcase fmt))

  ;; Set separator + quoting
  (cond
    ((= fmt "CSV")
     (setq sep "," quote T))
    (T
     (setq sep "\t" quote nil))
  )

  ;; Active document
  (setq doc (vla-get-ActiveDocument (vlax-get-acad-object)))

  ;; Output file
  (setq outname
        (strcat
          (getvar "DWGPREFIX")
          (vl-filename-base (getvar "DWGNAME"))
          (if (= fmt "CSV") ".csv" ".txt")
        )
  )

  ;; Open file
  (setq file (open outname "w"))
  (if (not file)
    (progn
      (princ "\nUnable to open output file.")
      (exit)
    )
  )

  ;; Select blocks with attributes
  (setq ss (ssget "_A" '((0 . "INSERT") (66 . 1))))

  (if ss
    (progn
      (setq attTags nil)

      (repeat (setq i (sslength ss))
        (setq ent (ssname ss (setq i (1- i))))
        (setq obj (vlax-ename->vla-object ent))

        ;; Filter block name
        (if (wcmatch (strcase (vla-get-EffectiveName obj)) "*00-UNIT*")
          (progn
            (setq atts (vlax-invoke obj 'GetAttributes))

            ;; Build header once
            (if (null attTags)
              (progn
                (setq attTags
                      (mapcar
                        '(lambda (a) (vla-get-TagString a))
                        atts
                      )
                )

                ;; SORT TAGS
                (setq attTags (vl-sort attTags '<))

                ;; Write header
                (write-line
                  (apply 'strcat
                    (append
                      (list (strcat "HANDLE" sep "BLOCKNAME" sep))
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

            ;; Build tag->value association list
            (setq attPairs
              (mapcar
                '(lambda (att)
                   (cons (strcase (vla-get-TagString att))
                         (vla-get-TextString att)))
                atts
              )
            )

            ;; Write row
            (write-line
              (apply 'strcat
                (append
                  (list
                    (strcat
                      (if quote "\"" "")
                      "'" (vla-get-Handle obj)
                      (if quote "\"" "")
                      sep
                    )
                    (strcat
                      (if quote "\"" "")
                      (vla-get-EffectiveName obj)
                      (if quote "\"" "")
                      sep
                    )
                  )
                  (mapcar
                    '(lambda (tag)
                       (setq val (cdr (assoc (strcase tag) attPairs)))
                       (strcat
                         (if quote "\"" "")
                         (if val val "")
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
      )
    )
  )

  (close file)

  (princ (strcat "\nAttributes exported to: " outname))
  (princ)
)