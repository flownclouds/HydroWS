(defun c:ezhbz ()
  (setvar "cmdecho" 0)
  (setq os (getvar "osmode"))
  (setvar "osmode" 0)
  (setq
    str_2 (strcat "\n给定文字高度<" (rtos (getvar "textsize") 2) ">:")
  )
  (setq th (getreal str_2))
  (if (= th nil)
    (setq th (getvar "textsize"))
  )
  (if (setq en (car (entsel "\n选择路径曲线:")))
    (redraw en 3)
    (*error* "\n没有选到曲线!")
  )
  (if (= nil (setq dist (getreal "\n给定间距<50m>:")))
    (setq dist 50.0)
  )
  (setq objen (vlax-ename->vla-object en))  
  (setq	objen_l	(vlax-curve-getDistAtPoint
		  objen
		  (vlax-curve-getEndPoint objen)
		)
  )
  (setq nn (fix (/ objen_l dist)))
  (repeat nn
    (setq pt1 (vlax-curve-getPointAtDist objen (* nn dist)))
    (setq zhz (* nn dist))
    (if	(< zhz 1000.0)
      (setq str_1 (strcat "K0+" (rtos zhz 2 3) "m"))
    )
    (if	(>= zhz 1000.0)
      (progn
	(setq nn1 (fix (/ zhz 1000.0)))
	(setq nn2 (- zhz (* 1000.0 nn1)))
	(setq str_1 (strcat "K" (rtos nn1 2) "+" (rtos nn2 2 3) "m"))
	(if (= nn2 0.0)
	  (setq str_1 (strcat "K" (rtos nn1 2) "+000m"))
	)
      )
    )
    (setq ang (a-get-Angle objen pt1))
    　　
    (setq pt2 (polar pt1 (+ (/ pi 2) ang) (* th 8)))
    　　
    (command "line" pt1 pt2 "")
    　　
    (setq ang2 (/ (* (angle pt2 pt1) 180) pi))
    (command "text" "bl" pt2 th ang2 str_1)
    　　
    (setq nn (1- nn))
  )

;;桩号标注文字与法线对齐的可以改这句：
;;(command "text" "bl" "non" pt2 th ang2 str_1)改为
;;(command "text" "M" "non" pt2 th ang2 str_1)
  　　
  ;; (redraw en 4)
  　　
  (setvar "osmode" os)
  　　
  (princ)
)
　　
;;切线角度pt-ang
　　
(defun a-get-Angle (ename point / p1 v1 pt-ang)
  　　
  (setq	v1		 (vlax-curve-getfirstderiv
			   　　
			   ename
			   　　
			   (vlax-curve-getparamatpoint ename point)
			   　　
			 )
	p1		 (mapcar '+ point v1)
	　　
			 pt-ang
	(angle point p1) 　　
  )
  　　
  pt-ang
  　　
)
