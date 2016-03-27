;此程序功能：将文字根据用户要求按左端、居中（垂直方向）、右、顶部、居中（水平方向）底部的方;式对齐，同时可调整其间距！文字有倾角亦是可适用的！
;说明：此次是第二次更新，经过测试未发现Bug；若各位坛友发现Bug，请反馈。

;;; ***文字对齐 程序开始***
(defun c:wzdq ()
  (princ
    "\n功能：将文字根据用户要求按左端、居中（垂直方向）、右、顶部、居中（水平方向）底部的方式对齐，同时可调整其间距及字高！\n"
  )
  (setvar "osmode" 15359)
  (setvar "cmdecho" 0)
  (if (not (setq ss (ssget '((0 . "TEXT,MTEXT")))))
    (progn (princ "\n未选中文字对象，程序退出。\n") (exit))
  )
  (command "undo" "be")
  (initget "L R M T B C")
  (if (not (setq kw
                  (getkword
                    "\n请选择对齐方式：[左端对齐(L)/以中心对齐（垂直方向）(M)/右端对齐(R)/顶部对齐(T)/以中心对齐（水平方向）(C)/底部对齐(B)/]<L>"
                  )
           )
      )
    (setq kw "L")
  )
  (initget "Y N")                        ;让用户选择是否调整文字之间的间距
  (if (not
        (setq kwGap
               (getkword "是否调整文字之间的间距？[是(Y)/否(N)]<Y>")
        )
      )
    (setq kwGap "Y")
  )
  (if (= kwGap "Y")
    (progn
      (initget 6)
      (if (not
            (setq
              gap (getdist "\n请指定排版后文字之间的间距：<3.0>")
            )
          )
        (setq gap 3.0)
      )
    )
  )
  (setq        i   0
        lst '()
  )
  (setvar "osmode" 0)
  (vl-load-com)
  (repeat (sslength ss)
    (setq txtentname (ssname ss i))
    (cond
      ((= kw "L")                        ;左端对齐
       (progn
         (command "_.justifytext" txtentname "" "ML")
         (wdy_wzdq_Ysort)
       )
      )
      ((= kw "M")                        ;以中心对齐（垂直方向）
       (progn
         (command "_.justifytext" txtentname "" "MC")
         (wdy_wzdq_Ysort)
       )
      )
      ((= kw "R")                        ;右端对齐
       (progn
         (command "_.justifytext" txtentname "" "MR")
         (wdy_wzdq_Ysort)
       )
      )
      ((= kw "T")                        ;顶部对齐
       (progn
         (command "_.justifytext" txtentname "" "TC")
         (wdy_wzdq_Xsort)
       )
      )
      ((= kw "B")                        ;底部对齐
       (progn
         (command "_.justifytext" txtentname "" "BC")
         (wdy_wzdq_Xsort)
       )
      )
      ((= kw "C")                        ;以中心对齐（水平方向）
       (progn
         (command "_.justifytext" txtentname "" "MC")
         (wdy_wzdq_Xsort)
       )
      )
    )
    (setq i (1+ i))
  )
;;;以重新排序后的表中的第一个文字对象作为参考对象
  (setq        entnam_base  (car (car lst))
        entdata_base (entget entnam_base)
        enttype_base (cdr (assoc 0 entdata_base))
  )
  (if (= enttype_base "TEXT")
    (setq tbox_base  (textbox (list (car entdata_base)))
          ptbl_base  (car tbox_base)
          pttr_base  (cadr tbox_base)
          pt_base    (cdr (assoc 11 entdata_base))
                                        ;读取文字对象的插入点
          ptx_base   (car pt_base)        ;插入点的X坐标
          pty_base   (cadr pt_base)        ;插入点的Y坐标
          ptx_pitch  ptx_base
          pty_pitch  pty_base
          heigh_base (cdr (assoc 40 entdata_base))
          width_base (abs (- (car pttr_base) (car ptbl_base)))
    )                                        ;若为单行文字
    (setq pt_base    (cdr (assoc 10 entdata_base))
                                        ;读取文字对象的插入点
          ptx_base   (car pt_base)        ;插入点的X坐标
          pty_base   (cadr pt_base)        ;插入点的Y坐标
          ptx_pitch  ptx_base
          pty_pitch  pty_base
          heigh_base (cdr (assoc 43 entdata_base))
                                        ;取多行文字的字体最大值
          width_base (cdr (assoc 42 entdata_base))
    )                                        ;若为多行文字
  )
  (setq i 1)
  (repeat (- (length lst) 1)
    (setq entnam_current  (car (nth i lst))
          entdata_current (entget entnam_current)
          enttype_current (cdr (assoc 0 entdata_current))
    )
    (if        (or (= kw "L") (= kw "R") (= kw "M")) ;左中右对齐时
      (progn (wdy_wzdq_type)
             (if (= kwGap "Y")
               (setq pty_pitch        (+ pty_pitch
                                   (* 0.5 heigh_base)
                                   (* 0.5 heigh_current)
                                   gap
                                )
                     heigh_base        heigh_current
               )                        ;若用户要求将文字间距设置为相同
               (setq pty_pitch (cadr pt_current))
                                        ;若用户未要求将文字间距设置为相同，即为原始值时
             )
             (setq pt (list ptx_base pty_pitch 0))
             (if (= enttype_current "TEXT")
               (entmod (subst (cons 11 pt)
                              (assoc 11 entdata_current)
                              entdata_current
                       )
               )
               (entmod (subst (cons 10 pt)
                              (assoc 10 entdata_current)
                              entdata_current
                       )
               )
             )
      )
    )
    (if        (or (= kw "T") (= kw "B") (= kw "C")) ;顶中底
      (progn (wdy_wzdq_type)
             (if (= kwGap "Y")
               (setq ptx_pitch        (+ ptx_pitch
                                   (* 0.5 width_base)
                                   (* 0.5 width_current)
                                   gap
                                )
                     width_base        width_current
               )                        ;若用户要求将文字间距设置为相同
               (setq ptx_pitch (car pt_current))
                                        ;若用户未要求将文字间距设置为相同，即为原始值时
             )
             (setq pt (list ptx_pitch pty_base 0))
             (if (= enttype_current "TEXT")
               (entmod (subst (cons 11 pt)
                              (assoc 11 entdata_current)
                              entdata_current
                       )
               )
               (entmod (subst (cons 10 pt)
                              (assoc 10 entdata_current)
                              entdata_current
                       )
               )
             )
      )
    )
    (setq i (1+ i))
  )
  (setvar "osmode" 15359)
  (command "undo" "e")
  (princ)
)

(defun wdy_wzdq_type ()
  (if (= enttype_current "TEXT")
    (setq tbox_current        (textbox (list (car entdata_current)))
          ptbl_current        (car tbox_current)
          pttr_current        (cadr tbox_current)
          pt_current        (cdr (assoc 11 entdata_current))
                                        ;读取文字对象的插入点
          heigh_current        (cdr (assoc 40 entdata_current))
          width_current        (abs (- (car pttr_current) (car ptbl_current)))
    )                                        ;若为单行文字
    (setq pt_current        (cdr (assoc 10 entdata_current))
                                        ;读取文字对象的插入点
          heigh_current        (cdr (assoc 43 entdata_current))
          width_current        (cdr (assoc 42 entdata_current))
    )                                        ;若为多行文字
  )
)
本帖隐藏的内容

;;;以X坐标比较进行排序
(defun wdy_wzdq_Xsort ()
  (setq
    inpoint (vlax-get (vlax-ename->vla-object txtentname)
                      'InsertionPoint
            )
  )
  (setq        lst (append
              (list (cons txtentname inpoint))
              lst
            )
  )
  (setq
    lst
     (vl-sort lst
              (function        (lambda        (e1 e2)
                          (if (equal (cadr e1) (cadr e2) 1e-5)
                            (if        (equal (caddr e1) (caddr e2) 1e-5)
                              (< (cadr e1) (cadr e2))
                              (< (caddr e1) (caddr e2))
                            )
                          )
                        )
              )
     )
  )
)
;;;以Y坐标比较进行排序
(defun wdy_wzdq_Ysort ()
  (setq
    inpoint (vlax-get (vlax-ename->vla-object txtentname)
                      'InsertionPoint
            )
  )
  (setq        lst (append
              (list (cons txtentname inpoint))
              lst
            )
  )
  (setq
    lst
     (vl-sort lst
              (function        (lambda        (e1 e2)
                          (if (equal (caddr e1) (caddr e2) 1e-5)
                            (if        (equal (cadr e1) (cadr e2) 1e-5)
                              (< (caddr e1) (caddr e2))
                              (< (cadr e1) (cadr e2))
                            )
                          )
                        )
              )
     )
  )
)