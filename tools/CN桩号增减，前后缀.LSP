(defun c:cn (/ se cn_1 st str sn0 sn1 st1 stp stm sts sg0 sgs ln len i j ei x)
  (command "color" "")
  (setq oldzin (getvar "dimzin"))
  (setvar "dimzin" 0)
  (princ "\n选取欲改变的序号或里程桩号: ")
  (setq se (ssget '((0 . "TEXT"))))
  (if se
    (progn
      (if (not cnd)
	(setq cnd 2)
      )
      (initget 1 "Pre Suf Kl")
      (setq cn_1
	     (getint "\n加前缀P/加后缀S/里程桩号增减K/<增加或减少的序数>: ")
      )
      (cond
	((= cn_1 "Pre")
	 (setq st (getstring "\n前缀字符串: "))
	 (cn_2)
	)
	((= cn_1 "Suf")
	 (setq st (getstring "\n后缀字符串: "))
	 (cn_2)
	)
	((= cn_1 "Kl") (cn_4))
	(T
	 (setq st cn_1)
	 (cn_2)
	)
      )
    )
    (princ "\n未选择序号字符串。")
  )
  (setvar "dimzin" oldzin)
  (princ)
)

(defun cn_2 ()				;选择集处理
  (setq	ln (sslength se)
	j  -1
  )
  (repeat ln
    (setq j (1+ j)
	  x (ssname se j)
    )
    (cn_3 x)
  )
)

(defun cn_3 (en)			;字符串处理
  (setq	ei  (entget en)
	str (cdr (assoc 1 ei))
  )
  (cond
    ((= cn_1 "Pre")
     (setq str (strcat st str)
	   ei  (subst (cons 1 str) (assoc 1 ei) ei)
     )
     (entmod ei)
    )
    ((= cn_1 "Suf")
     (setq str (strcat str st)
	   ei  (subst (cons 1 str) (assoc 1 ei) ei)
     )
     (entmod ei)
    )
    (T
     (setq i   1
	   len (strlen str)
	   sg0 0			;无后数字串
	   sgs 0			;无后缀
	   sn0 ""			;后数字串
	   stp ""			;前缀字符串
	   sts ""			;后缀字符串
     )
     (while (<= i len)
       (setq st1 (substr str 1 1)
	     str (substr str 2)
	     i	 (1+ i)
       )
       (if (and (>= (ascii st1) 48) (<= (ascii st1) 57)) ;数字判断
	 (if (= sgs 0)
	   (setq sn0 (strcat sn0 st1)	;数字串
		 sts ""
		 sg0 1
	   )
	   (setq stp (strcat stp sn0 sts)
					;有第二数字串时，将第一数字串作为前缀
		 sn0 st1
		 sts ""
		 sg0 1
		 sgs 0
	   )
	 )
	 (if (= sg0 0)
	   (setq stp (strcat stp st1))	;前缀串
	   (setq sts (strcat sts st1)	;后缀串
		 sgs 1
	   )
	 )
       )
     )
     (setq sn0 (itoa (+ (atoi sn0) st))
	   str (strcat stp sn0 sts)
	   ei  (subst (cons 1 str) (assoc 1 ei) ei)
     )
     (entmod ei)
    )
  )
)

(defun cn_5 ()
  (setq	ei  (entget en)
	str (cdr (assoc 1 ei))
  )
  (setq	i   1
	len (strlen str)
	sg0 0				;无后数字串
	sgs 0				;无后缀
	sn0 ""				;后数字串
	sn1 ""				;前数字串
	stp ""				;前缀字符串
	sts ""				;后缀字符串
	stm ""				;中字符
  )
  (while (<= i len)
    (setq st1 (substr str 1 1)
	  str (substr str 2)
	  i   (1+ i)
    )
    (if	(or (and (>= (ascii st1) 48) (<= (ascii st1) 57)) (= st1 "."))
					;数字判断
      (if (= sgs 0)
	(setq sn0 (strcat sn0 st1)	;数字串
	      sts ""
	      sg0 1
	)
	(if (or (= sts "+") (= sts "-"))
	  (setq	stp (strcat stp sn1 stm)
					;有第三数字串时，将第二数字串作为前缀
		sn1 sn0
		stm sts
		sn0 st1
		sts ""
		sg0 1
		sgs 0
	  )
	  (setq	stp (strcat stp sn0 sts)
		sn0 st1
		sg0 1
		sgs 0
	  )
	)
      )
      (if (= sg0 0)
	(setq stp (strcat stp st1))	;前缀串
	(setq sts (strcat sts st1)	;后缀串
	      sgs 1
	)
      )
    )
  )
  (if (/= stp "")
    (if (= (substr stp (setq len (strlen stp))) "-")
      (setq stp (substr stp 1 (- len 1))
	    sn1 (strcat "-" sn1)
      )
    )
  )
  (if (/= stm "")
    (setq sn0 ((eval (read stm)) (* (atoi sn1) 1000.0) (atof sn0)))
    (setq sn0 (atof sn0))
  )
  (setq sn0 (+ sn0 st) sn1 "0")
  (cond
    ((>= sn0 1000.0)
     (setq sn1 (itoa (fix (/ sn0 1000.0)))
	   sn0 (rem sn0 1000.0)
     )
    )
    ((< sn0 0.0)
     (setq sn1 (itoa (fix (/ sn0 1000.0))))
     (if (> (atoi sn1) 0)
       (setq sn0 (+ (rem sn0 1000.0) 1000.0))
       (setq sn0 (rem sn0 1000.0))
     )
    )
    (T (setq sn0 sn0 sn1 "0"))
  )
  (if (< sn0 0.0)
    (setq stm "-")
    (setq stm "+")
  )
  (if (< (abs sn0) 10.0)
    (setq sn0 (strcat "00" (rtos (abs sn0) 2 cnd)))
    (if	(< (abs sn0) 100.0)
      (setq sn0 (strcat "0" (rtos (abs sn0) 2 cnd)))
      (setq sn0 (rtos (abs sn0) 2 cnd))
    )
  )
  (setq	str (strcat stp sn1 stm sn0 sts)
	ei  (subst (cons 1 str) (assoc 1 ei) ei)
  )
  (entmod ei)
)



(defun cn_4 ()
  (initget 1 "D")
  (setq	st (getreal
	     (strcat "\n小数位数D( " (itoa cnd) " )/<增减桩号>: ")
	   )
  )
  (cond
    ((= st "D")
     (initget 4)
     (setq cnp (getint (strcat "\n保留小数位数< " (itoa cnd) " >: ")))
     (if cnp
       (setq cnd cnp)
     )
     (cn_4)
    )
    (T
     (setq ln (sslength se)
	   j  -1
     )
     (repeat ln
       (setq j	 (1+ j)
	     en  (ssname se j)
       )
       (cn_5)
     )
    )
  )
)