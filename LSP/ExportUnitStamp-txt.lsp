;;; ---------------------------------------------------------------------------
;;; UNITOUT - export the attributes of every *00-UNIT* block reference in
;;; ModelSpace to a tab separated <drawing name>.txt beside the drawing.
;;;
;;; Blocks are collected first so the header can be the union of every tag in
;;; use, then each row is written in header order and looked up by tag name.
;;; Taking the header from the first block and writing each block's attributes
;;; in its own order jumbles the columns as soon as two block definitions list
;;; the same tags in a different order, or one carries a tag another does not.
;;; ---------------------------------------------------------------------------

;; Join a list of strings, each followed by a tab (the layout Excel expects)
(defun UO:cat-fields (lst / s)
  (setq s "")
  (foreach f lst
    (setq s (strcat s f "\t"))
  )
  s
)

(defun c:UNITOUT (/ doc ms blk outname file atts attObj attTags attPairs
                    rows row tag)
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

  ;; -------------------------------------------------------------------------
  ;; Pass 1 - collect every unit block, and the union of the tags they use
  ;; -------------------------------------------------------------------------
  (setq attTags nil
        rows    nil
  )

  (vlax-for blk ms
    (if (and
          (= (vla-get-ObjectName blk) "AcDbBlockReference")
          (wcmatch (strcase (vla-get-EffectiveName blk)) "*00-UNIT*")
          ;; Compare against :vlax-true - :vlax-false is itself a non-nil
          ;; symbol, so testing the property on its own passes every block
          (= (vla-get-HasAttributes blk) :vlax-true)
        )
      (progn
        (setq atts     (vlax-invoke blk 'GetAttributes)
              attPairs nil
        )

        (foreach attObj atts
          ;; Tags are matched upper case, so a block defined with "Type" lands
          ;; in the same column as one defined with "TYPE"
          (setq tag (strcase (vla-get-TagString attObj)))

          ;; First occurrence wins if a block repeats a tag
          (if (not (assoc tag attPairs))
            (setq attPairs
                  (cons (cons tag (vla-get-TextString attObj)) attPairs)
            )
          )

          ;; Header keeps the order the tags are first met in
          (if (not (member tag attTags))
            (setq attTags (append attTags (list tag)))
          )
        )

        (setq rows
              (cons
                (list (vla-get-Handle blk)
                      (vla-get-EffectiveName blk)
                      attPairs
                )
                rows
              )
        )
      )
    )
  )

  (setq rows (reverse rows))

  ;; -------------------------------------------------------------------------
  ;; Pass 2 - write the header, then every row in header order
  ;; -------------------------------------------------------------------------
  (if (null rows)
    (princ "\nNo *00-UNIT* blocks with attributes found - nothing exported.")
    (progn
      (setq file (open outname "w"))
      (if (not file)
        (progn
          (princ "\nUnable to open output file.")
          (exit)
        )
      )

      (write-line
        (strcat "HANDLE\tBLOCKNAME\t" (UO:cat-fields attTags))
        file
      )

      (foreach row rows
        (setq attPairs (caddr row))
        (write-line
          (strcat
            ;; Leading apostrophe keeps the handle text once it reaches Excel
            "'" (car row) "\t"
            (cadr row) "\t"
            ;; A tag this block does not carry writes an empty field, so the
            ;; columns after it stay under their own header
            (UO:cat-fields
              (mapcar
                '(lambda (tag / val)
                   (setq val (cdr (assoc tag attPairs)))
                   (if val val "")
                 )
                attTags
              )
            )
          )
          file
        )
      )

      (close file)

      (princ
        (strcat
          "\n"
          (itoa (length rows))
          " unit blocks exported to: "
          outname
        )
      )
    )
  )

  (princ)
)

