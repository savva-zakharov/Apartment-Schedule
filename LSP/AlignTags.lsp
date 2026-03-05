(defun c:FixAngle1 ( / ss i ent obj rotW rotU ang props p ucsang)
  (vl-load-com)

  ;; angle of UCS X axis in WCS
  (setq ucsang (angle '(0 0 0) (trans '(1 0 0) 1 0)))

  (if (setq ss (ssget '((0 . "INSERT"))))
    (progn
      (setq i 0)
      (repeat (sslength ss)

        (setq ent (ssname ss i))
        (setq obj (vlax-ename->vla-object ent))

        ;; block rotation in WCS
        (setq rotW (vla-get-Rotation obj))

        ;; convert rotation relative to UCS
        (setq rotU (- rotW ucsang))

        ;; target dynamic property value
        (setq ang (- (/ pi 2) rotU))

        ;; update Angle1
        (setq props (vlax-invoke obj 'GetDynamicBlockProperties))

        (foreach p props
          (if (= (strcase (vla-get-PropertyName p)) "ANGLE1")
            (vla-put-Value p ang)
          )
        )

        (setq i (1+ i))
      )
    )
  )

  (princ)
)