;;--------------=={ Area to Attribute by Block Location }==-------------;;
;;                                                                      ;;
;;  This program populates a specified attribute on one or more        ;;
;;  attributed blocks with a Field Expression referencing the area of  ;;
;;  whichever selected closed polyline contains that block's           ;;
;;  insertion point (its "origin").                                    ;;
;;                                                                      ;;
;;  Upon issuing the command syntax 'A2BA' at the AutoCAD command-line, ;;
;;  the user is prompted to:                                           ;;
;;                                                                      ;;
;;    1. Select one or more CLOSED LWPOLYLINEs representing the areas  ;;
;;       (e.g. unit or room boundaries).                                ;;
;;    2. Select one or more attributed blocks (e.g. unit tag blocks)   ;;
;;       whose attribute should display the enclosing area.            ;;
;;    3. Enter the tag name of the attribute to populate. If left       ;;
;;       blank, "AREA" is assumed; if a block does not have a match    ;;
;;       for the given tag but has exactly one attribute, that single  ;;
;;       attribute is used instead.                                    ;;
;;                                                                      ;;
;;  For each selected block, the program tests its insertion point     ;;
;;  against every selected polyline. If the point lies within more     ;;
;;  than one polyline (e.g. nested boundaries), the SMALLEST enclosing ;;
;;  polyline is used, on the assumption that it is the most specific   ;;
;;  match. Blocks whose insertion point does not lie within any        ;;
;;  selected polyline, or which have no matching attribute, are        ;;
;;  reported and skipped.                                              ;;
;;                                                                      ;;
;;  The resulting Field Expression is identical in form to that        ;;
;;  produced by Lee Mac's Areas2Attribute (A2A) routine, and will      ;;
;;  update automatically if the polyline's area subsequently changes.  ;;
;;                                                                      ;;
;;  Notes / Limitations:                                                ;;
;;    - Only LWPOLYLINE entities are supported as area boundaries.      ;;
;;    - Containment testing is performed in the polyline's own XY       ;;
;;      plane (Z / elevation is ignored), which is appropriate for      ;;
;;      typical flat floor-plan geometry.                               ;;
;;    - Undo is grouped, so the entire operation can be reversed with   ;;
;;      a single U / CTRL+Z.                                            ;;
;;                                                                      ;;
;;----------------------------------------------------------------------;;
;;  Reuses the LM:objectid / LM:startundo / LM:endundo / LM:acdoc       ;;
;;  utility functions authored by Lee Mac (www.lee-mac.com), as also    ;;
;;  used by Areas2Attribute.lsp / Areas2Field.lsp in this folder.       ;;
;;----------------------------------------------------------------------;;
;;  Version 1.1    -    2026-08-24                                      ;;
;;----------------------------------------------------------------------;;

;; Remembers the last-used attribute tag for the remainder of the drawing
;; session (i.e. until AutoCAD is closed), so it can be offered as the
;; default the next time A2BA is run without needing to be retyped.
(if (not *a2ba-tag*) (setq *a2ba-tag* "AREA"))

(defun c:a2ba ( / *error* att atl blk doc fmt i n obj pl plines pnt sel1 sel2 tag best bestarea )

    (setq fmt "%lu2%qf1%pr0%ps[,]%ct8[1.000000000000000E-006]") ;; Field Formatting

    (defun *error* ( msg )
        (LM:endundo (LM:acdoc))
        (if (not (wcmatch (strcase msg t) "*break,*cancel*,*exit*"))
            (princ (strcat "\nError: " msg))
        )
        (princ)
    )

    (setq doc (LM:acdoc))

    (if
        (and
            (princ "\nSelect closed polylines defining areas: ")
            (setq sel1 (ssget '((0 . "LWPOLYLINE"))))
            (princ "\nSelect blocks to receive area attribute: ")
            (setq sel2 (ssget '((0 . "INSERT"))))
        )
        (progn
            (setq tag (strcase (LM:getstring-default (strcat "\nEnter attribute tag to populate <" *a2ba-tag* ">: ") *a2ba-tag*)))
            (setq *a2ba-tag* tag) ;; remember for the rest of the session

            ;; Build a cache of ( vla-object  point-list  area ) for every closed polyline
            (setq plines nil i (sslength sel1))
            (repeat i
                (setq obj (vlax-ename->vla-object (ssname sel1 (setq i (1- i)))))
                (if (= :vlax-true (vla-get-closed obj))
                    (setq plines (cons (list obj (LM:pline-points obj) (vla-get-area obj)) plines))
                    (princ (strcat "\nWarning: skipped open polyline (Handle " (vla-get-handle obj) ")."))
                )
            )

            (cond
                (   (null plines)
                    (princ "\nNo closed polylines were found in the selection - nothing to do.")
                )
                (   t
                    (LM:startundo doc)
                    (setq n 0 i (sslength sel2))
                    (repeat i
                        (setq blk (vlax-ename->vla-object (ssname sel2 (setq i (1- i)))))
                        (setq pnt (vlax-safearray->list (vlax-variant-value (vla-get-insertionpoint blk))))

                        ;; Find the smallest polyline whose boundary encloses the block's insertion point
                        (setq best nil bestarea nil)
                        (foreach pl plines
                            (if (and (LM:ptinpoly pnt (cadr pl))
                                     (or (null bestarea) (< (caddr pl) bestarea))
                                )
                                (setq best pl bestarea (caddr pl))
                            )
                        )

                        (if best
                            (progn
                                (setq atl (if (= :vlax-true (vla-get-hasattributes blk)) (vlax-invoke blk 'GetAttributes)) att nil)
                                (foreach a atl
                                    (if (= tag (strcase (vla-get-tagstring a)))
                                        (setq att a)
                                    )
                                )
                                (if (and (null att) (= 1 (length atl)))
                                    (setq att (car atl))
                                )
                                (if att
                                    (progn
                                        (vla-put-textstring att
                                            (strcat
                                                "%<\\AcObjProp Object(%<\\_ObjId "
                                                (LM:objectid (car best))
                                                ">%).Area \\f \"" fmt "\">%"
                                            )
                                        )
                                        (vl-cmdf "_.updatefield" (vlax-vla-object->ename att) "")
                                        (setq n (1+ n))
                                    )
                                    (princ (strcat "\nWarning: block (Handle " (vla-get-handle blk) ") has no attribute tagged \"" tag "\"."))
                                )
                            )
                            (princ (strcat "\nWarning: block (Handle " (vla-get-handle blk) ") origin does not lie within any selected polyline."))
                        )
                    )
                    (LM:endundo doc)
                    (princ (strcat "\n" (itoa n) " of " (itoa (sslength sel2)) " block attribute(s) updated."))
                )
            )
        )
    )
    (princ)
)

;; Get String with Default  -  helper
;; Prompts the user for a string, returning a supplied default if the user presses Enter.

(defun LM:getstring-default ( msg def / str )
    (setq str (getstring t msg))
    (if (or (null str) (= "" str)) def str)
)

;; Polyline Points  -  helper
;; Returns the list of (x y) vertex points of a supplied LWPOLYLINE vla-object,
;; in the polyline's own OCS/XY plane.

(defun LM:pline-points ( obj / crd lst )
    (setq crd (vlax-safearray->list (vlax-variant-value (vla-get-coordinates obj))))
    (while crd
        (setq lst (cons (list (car crd) (cadr crd)) lst)
              crd (cddr crd)
        )
    )
    (reverse lst)
)

;; Point in Polygon  -  helper (standard ray-casting / PNPOLY algorithm)
;; pt  - [lst] point to test, as (x y ...)
;; pts - [lst] list of polygon vertices, each (x y)
;; Returns T if pt lies within the polygon defined by pts, else nil.

(defun LM:ptinpoly ( pt pts / c i j n x xi xj y yi yj )
    (setq n (length pts) c nil i 0 j (1- n)
          x (car pt) y (cadr pt)
    )
    (while (< i n)
        (setq xi (car  (nth i pts)) yi (cadr (nth i pts))
              xj (car  (nth j pts)) yj (cadr (nth j pts))
        )
        (if (and (not (eq (> yi y) (> yj y)))
                 (< x (+ xi (/ (* (- xj xi) (- y yi)) (- yj yi))))
            )
            (setq c (not c))
        )
        (setq j i i (1+ i))
    )
    c
)

;; ObjectID  -  Lee Mac
;; Returns a string containing the ObjectID of a supplied VLA-Object
;; Compatible with 32-bit & 64-bit systems

(defun LM:objectid ( obj )
    (eval
        (list 'defun 'LM:objectid '( obj )
            (if (wcmatch (getenv "PROCESSOR_ARCHITECTURE") "*64*")
                (if (vlax-method-applicable-p (vla-get-utility (LM:acdoc)) 'getobjectidstring)
                    (list 'vla-getobjectidstring (vla-get-utility (LM:acdoc)) 'obj ':vlax-false)
                   '(LM:ename->objectid (vlax-vla-object->ename obj))
                )
               '(itoa (vla-get-objectid obj))
            )
        )
    )
    (LM:objectid obj)
)

;; Entity Name to ObjectID  -  Lee Mac
;; Returns the 32-bit or 64-bit ObjectID for a supplied entity name

(defun LM:ename->objectid ( ent )
    (LM:hex->decstr
        (setq ent (vl-string-right-trim ">" (vl-prin1-to-string ent))
              ent (substr ent (+ (vl-string-position 58 ent) 3))
        )
    )
)

;; Hex to Decimal String  -  Lee Mac
;; Returns the decimal representation of a supplied hexadecimal string

(defun LM:hex->decstr ( hex / foo bar )
    (defun foo ( lst rtn )
        (if lst
            (foo (cdr lst) (bar (- (car lst) (if (< 57 (car lst)) 55 48)) rtn))
            (apply 'strcat (mapcar 'itoa (reverse rtn)))
        )
    )
    (defun bar ( int lst )
        (if lst
            (if (or (< 0 (setq int (+ (* 16 (car lst)) int))) (cdr lst))
                (cons (rem int 10) (bar (/ int 10) (cdr lst)))
            )
            (bar int '(0))
        )
    )
    (foo (vl-string->list (strcase hex)) nil)
)

;; Start Undo  -  Lee Mac
;; Opens an Undo Group.

(defun LM:startundo ( doc )
    (LM:endundo doc)
    (vla-startundomark doc)
)

;; End Undo  -  Lee Mac
;; Closes an Undo Group.

(defun LM:endundo ( doc )
    (while (= 8 (logand 8 (getvar 'undoctl)))
        (vla-endundomark doc)
    )
)

;; Active Document  -  Lee Mac
;; Returns the VLA Active Document Object

(defun LM:acdoc nil
    (eval (list 'defun 'LM:acdoc 'nil (vla-get-activedocument (vlax-get-acad-object))))
    (LM:acdoc)
)

;;----------------------------------------------------------------------;;

(vl-load-com)
(princ "\n:: Areas2AttributeByBlockLocation.lsp | Version 1.0 :: Type \"A2BA\" to Invoke ::")
(princ)

;;----------------------------------------------------------------------;;
;;                             End of File                              ;;
;;----------------------------------------------------------------------;;
