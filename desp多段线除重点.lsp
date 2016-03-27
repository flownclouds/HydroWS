(prompt "\n●***多段线除重点***●\n◎§※命令：DESP※§◎")
(defun c:DESP( / l Sel data newdata en enp js_n)
  (Setvar "Cmdecho" 0) (Prompt "\n\r★☆★选择需要处理的多段线：")
  (SetQ Sel (SsGet (list(cons 0 "lwpolyline")))
	L (SsLength Sel) ;;获取对象
	m 0 js_n 0
	) (Repeat L
	    (SetQ en (SsName Sel m)
		  data(entget en)
		  n 0
		  enp_js t
		  newdata NIL
		  js nil
		  )
	    (while
	      enp_js
	      (setq enp(nth n data)) ;;对组码进行循环，找出重复点
	      (if(and (member enp newdata)(= 10(car enp)))
		(progn (setq n (+ n 3) js T) )
		(setq newdata (cons enp newdata)) ;;筛选组码，去重点
		)
	      (setq n (1+ n))
	      (setq enp_js(nth n data))
	      )
	    (setq newdata(reverse newdata)
		  m (1+ m)
		  )
	    (entmod newdata) ;;更新图元，得到没有重复点的多段线
	    (if js
	      (setq js_n (1+ js_n))
	      )
	    )
  (Setvar "Cmdecho" 1)
  (princ "\n▲◆§共处理多段线条数=")
  js_n
  )