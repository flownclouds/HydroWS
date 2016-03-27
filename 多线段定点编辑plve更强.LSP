(defun c:plve (/ os en entl pt ptend s pts inline)
  (princ
    "\n plve====lwpolyline vert edit,增加/删除pl线顶点--vb1.0--lxx.2001.7"
  )
  (princ "\nplve.多义线顶点增删程序--梁雄啸.2001")
  (command "undo" "be")
  (setvar "cmdecho" 0)
  (princ "\n当取点为 pl 线的顶点时---删除/否---增加顶点")
  (setq	os  (getvar "osmode")
	en  (entsel "\n选择要增/删顶点的多义线:")
	ent (car en)
  )
  (setvar "osmode" 551)
  (initget 128)
  (princ "\n取点为 pl 线的顶点时==>删除 / 否==>增加顶点")


  (if (or (not en)(not(wcmatch  (cdr(assoc 0(entget (car en))))"*POLYLINE")))
      (progn (setvar "osmode" os)
	     (command "undo" "e")
	     (exit)
      )
  )
  (setq	s (ssadd)
	s (ssadd (car en) s)
  )
  (sssetfirst s s)
  


(while
    (setq pt (getpoint "\n自动识别增&删--取点/enter-退出:"))					
;;;    (command "line" "non" pt "non" "0,0,0" "")
;下面这句不同，导致错误
    (setq pt1 (vlax-curve-getClosestPointTo
	       (vlax-ename->vla-object ent)
	       pt
	     )
    )
    (if (<(distance pt1 pt)1e-4)(setq inline t))
    (setq pt pt1)(print inline) 
    
    
     (setq entl (entget ent))    
     (foreach n	entl
       (if (= 10 (car n))
	 (setq pts (cons (cdr n) pts))
       )
     )
     (if (member (list (car pt) (cadr pt)) pts)
       (eplv)				
       (aplv)				
     )
     (sssetfirst s s)
  )		
  (setq s (ssget))
  (setvar "osmode" os)
  (command "undo" "e")
  (princ)
)
;;;;删除顶点
(defun eplv (/ ptl entl2)
  (princ "\nok")
  (setq	ptl   (cons 10 (list (car pt) (cadr pt))) 
	entl2 '()
  )
  (foreach n entl
    (if	(not (equal n ptl 1e-4));这里不同
      (setq entl2 (cons n entl2))
    )
  )
  (setq entl (reverse entl2))
  (entmod entl)
)
;;;;增加顶点
(defun aplv ()
  (initget 129)
  (if inline  (setq npt (getpoint pt "\n 新的定位点<不做变动>:"))(setq npt ""))
  (if
    (= (type pt) 'LIST)
     (progn
       (print pt)
       (print npt)
       (command "_.break" ent pt "@" )
       (command "_.pedit" ent "j" (entlast) "" "")
       (while  (= (type npt) 'LIST)
	 (setq entl  (entget ent)
	       ptl   (cons 10 (list (car pt) (cadr pt)))
	       nptl  (cons 10 (list (car npt) (cadr npt)))
	       entl2 '()
	 )
	 (foreach n entl
	   (if (equal n ptl 1e-2)
	     (setq entl2 (cons nptl entl2))
	     (setq entl2 (cons n entl2))
	   )
	 )
	 (setq entl (reverse entl2)
	       npt  nil
	 )
	 (entmod entl)
       )
     )
  )
)
