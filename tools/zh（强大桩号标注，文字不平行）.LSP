;;这个程序功能很强大，但是在运行时提示有错误的地方cad有如下提示“命令: ZH
;;比例尺1:<1000>: 100
;;前缀<无>:000+123
;;保留小数位数<3>:
;;当前图形中的字所有体样式名列表:
;;    SILIN,ANNOTATIVE,STANDARD,
;;当前使用的字体样式:silin，当前字高:0.20
;;输入字体样式名<silin>:
;;指定字高<0.20>:3
;;请选择要标注桩号的多义线：AcDbPolyline
;;标注方向--铅直(Vertical)/半交角(Halfangle)/自由(Freedom)<Halfangle>:v
;;请选择桩号为0＋000.000的点<线起点>: <对象捕捉 开>
;;请点取标注位置：; 错误: no function definition: M_DRAWTEXT”

;;不知道该如何更改？烦请帮个忙！程序如下：
;;——标注桩号——
(vl-load-com)
(defun c:zh (/   m_oldcmdecho   m_oldosmode
      m_data  m_databl    m_dataqz  m_datajd
      m_datazt  m_datazt1   m_datazg  m_datazg1
      style_list  style_flag  m_kw  m_flag
      m_zt  m_zg      m_bzx  m_vlaobj
      m_flag  m_ptstart   m_pt  m_ss
      m_startlength      m_zhlength  m_alignment
      m_rotate  m_kw      m_ptzh  m_wz
     )
  (setq m_oldcmdecho (getvar "cmdecho"))
  (setvar "cmdecho" 0)
  (setq m_oldosmode (getvar "osmode"))
  (setvar "osmode" 16385)
  (initget (+ 2 4))
  (if (setq m_databl (getint "\n比例尺1:<1000>: "))
    (setq m_databl (itoa m_databl))
    (setq m_databl "1000")
  )
  (setq m_dataqz (getstring "\n前缀<无>:"))
  (initget (+ 2 4))
  (if (setq m_datajd (getint "\n保留小数位数<3>:"))
    (setq m_datajd (itoa m_datajd))
    (setq m_datajd "3")
  )
  (setq m_datazt (getvar "TEXTSTYLE")) ;当前字体样式名
  (setq m_datazg (getvar "TEXTSIZE")) ;当前字高
  (setq style_list (cons (cdadr (tblnext "style" T)) '()))
  (while (setq style_flag (tblnext "style"))
    (setq style_list (cons (cdadr style_flag) style_list))
  )     ;读出当前图中的所有文字样式
  (setq style_list (vl-remove "" style_list)) ;去掉空文字样式
  (setq style_list (mapcar 'strcase style_list))
  (princ
    (strcat "\n当前图形中的字所有体样式名列表: \n    "
     (apply 'strcat
     (mapcar '(lambda (x) (strcat x ",")) style_list)
     )
    )
  )
  (princ
    (strcat "\n当前使用的字体样式:"
     m_datazt
     "，当前字高:"
     (rtos m_datazg 2 2)
    )
  )
  (setq m_flag 1)
  (while m_flag
    (if (/= ""
     (setq m_datazt1
     (getstring (strcat "\n输入字体样式名<" m_datazt ">:")
     )
     )
)
      (if (not (member (strcase m_datazt1) style_list))
(princ (strcat "  --->错误: 当前图形中未找到字体样式:"
         (strcase m_datazt1)
        )
)
(progn
   (setq m_datazt m_datazt1)
   (setq m_flag nil)
)
      )
      (setq m_flag nil)
    )
  )
  (initget (+ 2 4))
  (if (setq m_datazg1
      (getreal (strcat "\n指定字高<" (rtos m_datazg 2 2) ">:")
      )
      )
    (setq m_datazg m_datazg1)
  )
  (setq m_data (list (list m_databl m_dataqz m_datajd "" "" "")
       (list m_datazt m_datazg)
        )
  )
  (setq m_zt (nth 0 (cadr m_data)))
  (setq m_zg (nth 1 (cadr m_data)))
  (command "undo" "be")
  (if (setq m_bzx (entsel "\n请选择要标注桩号的多义线："))
    (progn
      (setq m_bzx (car m_bzx))
      (setq m_vlaobj (vlax-ename->vla-object m_bzx))
      (princ (vla-get-objectname m_vlaobj)) ;显示实体名
      (if (or (= "AcDbPolyline" (vla-get-objectname m_vlaobj))
       (= "AcDb2DPolyline" (vla-get-objectname m_vlaobj))
   )
(progn
   (initget "Vertical Halfangle Freedom")
   (if (not (setq m_kw
     (getkword
       "\n标注方向--铅直(Vertical)/半交角(Halfangle)/自由(Freedom)<Halfangle>:"
     )
     )
       )
     (setq m_kw "Halfangle") ;默认标注方向为半交角
   )
   (cond ((= "Vertical" m_kw) (m_bzzh_Vertical m_data m_bzx))
  ((= "Halfangle" m_kw) (m_bzzh_Halfangle m_data m_bzx))
  ((= "Freedom" m_kw) (m_bzzh_Freedom m_data m_bzx))
   )
)
(princ "——>选择实体不是多段线！")
      )
    )
  )
  (command "undo" "e")
  (setvar "osmode" m_oldosmode)
  (setvar "cmdecho" m_oldcmdecho)
  (princ)
)
(defun m_bzzh_vertical (m_data     m_bzx /     m_flag
   m_vlaobj    m_ptstart m_pt
   m_startlength  m_ss     m_zxlength
   m_alignment m_rotate m_kw     m_zg
   m_zt     m_wzang m_ptzh     m_oldosmode
         )
  (setq m_zt (nth 0 (cadr m_data)))
  (setq m_zg (nth 1 (cadr m_data)))
  (setq m_vlaobj (vlax-ename->vla-object m_bzx))
  (setq m_flag T)
  (while m_flag    ;设定退出标志
    (if (setq
   m_ptstart (getpoint "\n请选择桩号为0＋000.000的点<线起点>:")
)
      (progn
(setq
   m_pt (vlax-curve-getclosestpointto m_vlaobj m_ptstart nil)
)
(if (> (distance m_pt m_ptstart) 0.0)
   (princ "\n所选点不在曲线上！")
   (progn
     (setq m_ptstart m_pt)
     (setq m_flag nil)  ;退出循环
   )
)
      )
      (progn
(setq m_ptstart (vlax-curve-getstartpoint m_vlaobj))
     ;默认多义线起点
(setq m_flag nil)  ;退出循环
      )
    )
  )
  (setq m_startlength (vlax-curve-getdistatpoint m_vlaobj m_ptstart))
     ;取得开始点距多义线起点点的长度
  (princ "\n请点取标注位置：")
  (setq m_ss (ssadd))
  (while (/= 3 (car (setq m_pt (grread T 5))))
    (if m_ss
      (command "erase" m_ss "")
    )
    (setq m_zxlength (distance m_ptstart (cadr m_pt)))
    (if (> (car (cadr m_pt)) (car m_ptstart))
      (progn
(setq m_alignment "mr")
(setq m_rotate (angtof "270"))
(setq m_wzang (angtof "90"))
      )
      (progn
(setq m_alignment "ml")
(setq m_rotate (angtof "90"))
(setq m_wzang (- 0 (angtof "90")))
      )
    )
    (setq m_pt (polar m_ptstart (angtof "270") m_zxlength))
    (setq m_oldosmode (getvar "osmode"))
    (setvar "osmode" 16385)
    (command "pline" m_ptstart m_pt "")
    (setvar "osmode" m_oldosmode)
    (ssadd (entlast) m_ss)
    (setq
      m_pt (polar m_pt (+ (angtof "270") m_wzang) (* (/ 5 6.0) m_zg))
    )
    (m_DrawText m_zt m_alignment m_pt m_zg m_rotate "0+000.000")
    (ssadd (entlast) m_ss)
  )
  (initget "Selectpoint Allvertex Insertpoint")
  (if (not (setq m_kw
    (getkword
      "\n选点标注(Selectpoint)/全部顶点标注(Allvertex)/插入标注(Insertpoint)<Allvertex>:"
    )
    )
      )
    (setq m_kw "Allvertex")  ;默认全部顶点标注
  )
  (print m_kw)
  (cond ((= "Selectpoint" m_kw)  ;点选标注(Selectpoint)
  (while (setq m_pt (getpoint "\n桩号点<退出>:"))
    (setq m_ptzh (vlax-curve-getclosestpointto m_vlaobj m_pt))
    (if (> (distance m_pt m_ptzh) 0.0)
      (princ "所选点不在曲线上!")
      (m_bzzh1_1 (list m_ptzh (angtof "270") m_zxlength)
   (list m_wzang m_alignment m_rotate)
   m_vlaobj
   m_startlength
   m_data
      )
    )
  )
)
((= "Allvertex" m_kw)  ;全部顶点标注(Vertext)
  (m_bzzh1_2 m_zt  m_zg       m_zxlength   m_wzang
      m_alignment  m_rotate     m_vlaobj    m_startlength
      m_data  m_ptstart
     )
)
((= "Insertpoint" m_kw)  ;插入标注(Insertpoint)
  (m_bzzh1_2 m_zt  m_zg       m_zxlength   m_wzang
      m_alignment  m_rotate     m_vlaobj    m_startlength
      m_data  m_ptstart
     )
  (while (setq m_pt (getpoint "\n桩号点<退出>:"))
    (setq m_ptzh (vlax-curve-getclosestpointto m_vlaobj m_pt))
    (if (> (distance m_pt m_ptzh) 0.0)
      (princ "所选点不在曲线上!")
      (m_bzzh1_1 (list m_pt (angtof "270") m_zxlength)
   (list m_wzang m_alignment m_rotate)
   m_vlaobj
   m_startlength
   m_data
      )
    )
  )
)
  )
)
(defun m_bzzh1_1 (m_zxlist     m_wzlist     m_vlaobj  m_startlength
    m_data       /     m_oldosmode  m_zt
    m_zg        m_bl     m_qz  m_xsws
    m_wz        m_zxpt     m_zxrotate  m_zxlength
    m_wzpt       m_wzang     m_wzalignment
    m_wzrotate
   )
  ;;m_zhlist—>(要标注的桩线点m_zxpt 桩线旋转角m_zxrotate 桩线长度m_zxlength)
  ;;m_wzlist—>(文字对齐方式m_wzalignment 文字旋转角m_wzrotate)
  (setq m_zt (nth 0 (cadr m_data))) ;字体
  (setq m_zg (nth 1 (cadr m_data))) ;文字高度
  (setq m_bl (read (nth 0 (car m_data)))) ;图纸比例
  (setq m_qz (nth 1 (car m_data))) ;前缀
  (setq m_xsws (read (nth 2 (car m_data)))) ;小数位数
  (setq m_zxpt (nth 0 m_zxlist))
  (setq m_zxrotate (nth 1 m_zxlist))
  (setq m_zxlength (nth 2 m_zxlist))
  (setq m_wzang (nth 0 m_wzlist))
  (setq m_wzalignment (nth 1 m_wzlist))
  (setq m_wzrotate (nth 2 m_wzlist))
  (setq m_wzpt (polar m_zxpt m_zxrotate m_zxlength))
  (setq m_oldosmode (getvar "osmode"))
  (setvar "osmode" 16385)
  (command "pline" m_zxpt m_wzpt "")
  (setvar "osmode" m_oldosmode)
  (setq m_wzpt (polar m_wzpt (+ m_zxrotate m_wzang) (* (/ 5 6.0) m_zg)))
  (setq m_wz (- (vlax-curve-getdistatpoint m_vlaobj m_zxpt)
  m_startlength
      )
  )
  (setq m_wz (/ (* (/ m_bl 1000.0) m_wz) 1000.0))
     ;###.############...(km)
  (if (> m_wz 0.0)
    (progn
      (setq m_wz
      (strcat
        m_qz
        (itoa (fix m_wz))
        "+"
        (substr
   (rtos m_wz 2 4)
   (+ 2 (vl-string-position (ascii ".") (rtos m_wz 2 4)))
   3
        )
        "."
        (substr
   (rtos m_wz 2 (+ 3 m_xsws))
   (+ 5
      (vl-string-position (ascii ".") (rtos m_wz 2 (+ 3 m_xsws)))
   )
   m_xsws
        )
      )
      )
    )
    (progn
      (setq m_wz (abs m_wz))
      (setq m_wz
      (strcat
        m_qz
        (itoa (fix m_wz))
        "-"
        (substr
   (rtos m_wz 2 4)
   (+ 2 (vl-string-position (ascii ".") (rtos m_wz 2 4)))
   3
        )
        "."
        (substr
   (rtos m_wz 2 (+ 3 m_xsws))
   (+ 5
      (vl-string-position (ascii ".") (rtos m_wz 2 (+ 3 m_xsws)))
   )
   m_xsws
        )
      )
      )
    )
  )
  (m_DrawText m_zt m_wzalignment m_wzpt m_zg m_wzrotate m_wz)
)
(defun m_bzzh1_2 (m_zt       m_zg   m_zxlength  m_wzang
    m_alignment m_rotate   m_vlaobj    m_startlength
    m_data      m_ptstart   /       m_ptlist
    m_pclist    n    m_pt
   )
  (setq m_ptlist (vla-get-coordinates m_vlaobj))
  (setq m_ptlist (vlax-safearray->list (vlax-variant-value m_ptlist)))
  (setq n 0)
  (repeat (/ (length m_ptlist) 2)
    (setq m_pclist (cons (list (nth n m_ptlist) (nth (1+ n) m_ptlist))
    m_pclist
     )
    )
    (setq n (+ 2 n))
  )     ;求出多义线的全部顶点
  (repeat (length m_pclist)
    (setq m_pt (car m_pclist))
    (if (> (distance (list (car m_ptstart) (cadr m_ptstart)) m_pt)
    0.000001
)
      (m_bzzh1_1 (list m_pt (angtof "270") m_zxlength)
   (list m_wzang m_alignment m_rotate)
   m_vlaobj
   m_startlength
   m_data
      )
    )
    (setq m_pclist (cdr m_pclist))
  )
)
(defun m_bzzh_Halfangle (m_data     m_bzx      /   m_zt
    m_zg     m_vlaobj   m_flag   m_ptstart
    m_pt     m_stratlength   m_ptlist
    m_ss     m_zxlength m_ptnearst m_angle
    m_angle    m_pt1      m_pt2   m_wzang
    m_alignment        m_rotate   m_olsosmode
    m_kw     m_zxpt     m_qz   m_wz
    m_xsws
   )
  (princ "半交角标注！")
  (setq m_zt (nth 0 (cadr m_data)))
  (setq m_zg (nth 1 (cadr m_data)))
  (setq m_qz (nth 1 (car m_data))) ;前缀
  (setq m_xsws (read (nth 2 (car m_data)))) ;小数位数
  (setq m_vlaobj (vlax-ename->vla-object m_bzx))
  (setq m_wz
  (strcat m_qz
   "0+000."
   (substr "000000000000000000000000000000000000000000000"
    1
    m_xsws
   )
  )
  )
  (setq m_flag T)
  (while m_flag    ;设定退出标志
    (if (setq m_ptstart
        (getpoint (strcat "\n请选择桩号为" m_wz "的点<多段线起点>:")
        )
)
      (progn
(setq
   m_pt (vlax-curve-getclosestpointto m_vlaobj m_ptstart)
)
(if (> (distance m_pt m_ptstart) 0.0)
   (princ "\n所选点不在曲线上！")
   (progn
     (setq m_ptstart m_pt)
     (setq m_flag nil)  ;退出循环
   )
)
      )
      (progn
(setq m_ptstart (vlax-curve-getstartpoint m_vlaobj))
     ;默认多义线起点
(setq m_flag nil)  ;退出循环
      )
    )
  )
  (setq m_startlength (vlax-curve-getdistatpoint m_vlaobj m_ptstart))
     ;取得开始点距多义线起点点的长度
  (setq m_ptlist (m_searchhd1 m_bzx)) ;返回带弧段圆心的多段线顶点表
  (setq m_ptnearst (m_searchpt m_ptstart m_vlaobj m_ptlist))
     ;返回两个端点
  (setq m_angle (m_halfangle m_ptnearst m_ptstart)) ;返回角度
  (princ "\n请点取标注位置：")
  (setq m_oldosmode (getvar "osmode"))
  (setvar "osmode" 16385)
  (setq m_ptstart (list (car m_ptstart) (cadr m_ptstart) 0.0))
     ;去掉z坐标值
  (setq m_ss (ssadd))
  (while (/= 3 (car (setq m_pt (grread T 5))))
    (if m_ss
      (command "erase" m_ss "")
    )
    (setq m_zxlength (distance m_ptstart (cadr m_pt))) ;桩线长度
    (setq m_pt1 (polar m_ptstart m_angle m_zxlength))
    (setq m_pt2 (polar m_ptstart (+ m_angle (angtof "180")) m_zxlength))
    (if (> (distance m_pt1 (cadr m_pt))
    (distance m_pt2 (cadr m_pt))
)
      (setq m_pt1 m_pt2)
    )
    (command "pline" m_ptstart m_pt1 "") ;绘制桩号线
    (ssadd (entlast) m_ss)
    (setq m_angle (angle m_ptstart m_pt1))
    (if (and (> m_angle (angtof "90")) (<= m_angle (angtof "270")))
     ;在二、三象限
      (progn    ;在二、三象限
(setq m_alignment "ml")
(setq m_rotate (+ m_angle (angtof "180")))
(setq m_wzang (- 0 (angtof "90"))) ;负90度
      )
      (progn    ;在一、四象限
(setq m_alignment "mr")
(setq m_rotate m_angle)
(setq m_wzang (angtof "90")) ;正90度
      )
    )
    (setq m_pt (polar m_pt1 (+ m_angle m_wzang) (* (/ 5 6.0) m_zg)))
    (m_DrawText m_zt m_alignment m_pt m_zg m_rotate m_wz)
    (ssadd (entlast) m_ss)
  )
  (setvar "osmode" m_oldosmode)
  (initget "Selectpoint Allvertex Divide Insertpoint Ctrol")
  (if (not (setq m_kw
    (getkword
      "\n选点(Selectpoint)/全部顶点(Allvertex)/插入(Insertpoint)/等距(Divide)/手控(Ctrol)<Ctrol>:"
    )
    )
      )
    (setq m_kw "Ctrol")   ;默认全部顶点标注
  )
  (print m_kw)
  (cond ((= "Selectpoint" m_kw)  ;点选标注(Selectpoint)
  (m_bzzh1_halfangle_Selectpoint
    m_zxlength m_startlength m_ptlist m_vlaobj m_data)
)
((= "Allvertex" m_kw)  ;全部顶点标注(Vertext)
  (m_bzzh1_halfangle_Allvertex
    m_zxlength m_startlength m_ptlist m_vlaobj m_data)
)
((= "Insertpoint" m_kw)  ;插入标注(Insertpoint)
  (m_bzzh1_halfangle_Allvertex
    m_zxlength m_startlength m_ptlist m_vlaobj m_data)
  (m_bzzh1_halfangle_Selectpoint
    m_zxlength m_startlength m_ptlist m_vlaobj m_data)
)
((= "Divide" m_kw)  ;等间距标注(Divide)
  (m_bzzh1_halfangle_Divide
    m_zxlength m_startlength m_ptlist m_vlaobj m_data)
)
((= "Ctrol" m_kw)  ;手控模式标注(Ctrol);控制标注线和文字位置
  (m_bzzh1_halfangle_Ctrol
    m_zxlength m_startlength m_ptlist m_vlaobj m_data)
)
  )
)
(defun m_bzzh1_halfangle_Ctrol (m_zxlength m_startlength
    m_ptlist   m_vlaobj   m_data
    /    m_zt       m_zg
    m_bl    m_pt       m_wz
    m_zxpt    m_ptnearst m_angle
    m_ss    m_oldosmode
    m_pt1    m_pt2      m_wzang
    m_alignment       m_rotate
    m_qz    m_sxws
          )
  (setq m_zt (nth 0 (cadr m_data))) ;字体
  (setq m_zg (nth 1 (cadr m_data))) ;文字高度
  (setq m_bl (read (nth 0 (car m_data)))) ;图纸比例
  (setq m_qz (nth 1 (car m_data))) ;前缀
  (setq m_xsws (read (nth 2 (car m_data)))) ;小数位数
  (while (setq m_pt (getpoint "\n桩号点<退出>:"))
    (setq m_zxpt (vlax-curve-getclosestpointto m_vlaobj m_pt))
    (if (> (distance m_pt m_zxpt) 0.0)
      (princ "所选点不在曲线上!")
      (progn
(setq m_ptnearst (m_searchpt m_zxpt m_vlaobj m_ptlist))
     ;返回两个端点
(setq m_angle (m_halfangle m_ptnearst m_zxpt)) ;返回角度
(setq m_oldosmode (getvar "osmode"))
(setvar "osmode" 16385)
(setq m_ss (ssadd))
(while (/= 3 (car (setq m_pt (grread T 5))))
   (if m_ss
     (command "erase" m_ss "")
   )
   (setq m_pt1 (polar m_zxpt m_angle m_zxlength))
   (setq
     m_pt2 (polar m_zxpt (+ m_angle (angtof "180")) m_zxlength)
   )
   (if (> (distance m_pt1 (cadr m_pt))
   (distance m_pt2 (cadr m_pt))
       )
     (setq m_pt1 m_pt2)
   )
   (command "pline" m_zxpt m_pt1 "") ;绘制桩号线
   (ssadd (entlast) m_ss)
   (setq m_angle (angle m_zxpt m_pt1))
   (if
     (and (> m_angle (angtof "90")) (<= m_angle (angtof "270")))
     ;在二、三象限
      (progn   ;在二、三象限
        (setq m_alignment "ml")
        (setq m_rotate (+ m_angle (angtof "180")))
      )
      (progn   ;在一、四象限
        (setq m_alignment "mr")
        (setq m_rotate m_angle)
      )
   )
   (if (> (angle m_zxpt (cadr m_pt)) m_angle)
     (setq m_wzang (angtof "90")) ;正90度
     (setq m_wzang (- 0 (angtof "90"))) ;负90度
   )
   (setq
     m_pt (polar m_pt1 (+ m_angle m_wzang) (* (/ 5 6.0) m_zg))
   )
   (setq m_wz (- (vlax-curve-getdistatpoint m_vlaobj m_zxpt)
   m_startlength
       )
   )
   (setq m_wz (/ (* (/ m_bl 1000.0) m_wz) 1000.0))
     ;###.############...(km)
   (if (> m_wz 0.0)
     (progn
       (setq m_wz (strcat
      m_qz
      (itoa (fix m_wz))
      "+"
      (substr (rtos m_wz 2 4)
       (+ 2
          (vl-string-position
     (ascii ".")
     (rtos m_wz 2 4)
          )
       )
       3
      )
      "."
      (substr (rtos m_wz 2 (+ 3 m_xsws))
       (+ 5
          (vl-string-position
     (ascii ".")
     (rtos m_wz 2 (+ 3 m_xsws))
          )
       )
       m_xsws
      )
    )
       )
     )
     (progn
       (setq m_wz (abs m_wz))
       (setq m_wz (strcat
      m_qz
      (itoa (fix m_wz))
      "-"
      (substr (rtos m_wz 2 4)
       (+ 2
          (vl-string-position
     (ascii ".")
     (rtos m_wz 2 4)
          )
       )
       3
      )
      "."
      (substr (rtos m_wz 2 (+ 3 m_xsws))
       (+ 5
          (vl-string-position
     (ascii ".")
     (rtos m_wz 2 (+ 3 m_xsws))
          )
       )
       m_xsws
      )
    )
       )
     )
   )
   (m_DrawText m_zt m_alignment m_pt m_zg m_rotate m_wz)
   (ssadd (entlast) m_ss)
)
(setvar "osmode" m_oldosmode)
      )
    )
  )
)
(defun m_bzzh1_halfangle_Divide (m_zxlength m_startlength
     m_ptlist   m_vlaobj   m_data
     /     m_zt       m_zg
     m_bl     m_qz       m_xsws
     m_jj     i        n
     m_zxpt     m_ptnearst m_angle
     m_alignment        m_rotate
     m_wzang
    )
  (setq m_zt (nth 0 (cadr m_data))) ;字体
  (setq m_zg (nth 1 (cadr m_data))) ;文字高度
  (setq m_bl (read (nth 0 (car m_data)))) ;图纸比例
  (setq m_qz (nth 1 (car m_data))) ;前缀
  (setq m_xsws (read (nth 2 (car m_data)))) ;小数位数
  (initget (+ 2 4))   ;阻止0值和负值
  (setq m_jj (getint "\n请输入间距(m)<50>:"))
  (if (not m_jj)
    (setq m_jj 50)
  )     ;默认间距50m
  (setq n (fix (/ (- (vlax-curve-getdistatpoint
         m_vlaobj
         (vlax-curve-getendpoint m_vlaobj)
       )
       m_startlength
    )
    (/ (* m_jj 1000.0) m_bl)
        )
   )
  )
  (setq i 1)
  (repeat n
    (setq m_zxpt (vlax-curve-getpointatdist
     m_vlaobj
     (+ m_startlength (/ (* i m_jj 1000.0) m_bl))
   )
    )
    (setq m_ptnearst (m_searchpt m_zxpt m_vlaobj m_ptlist))
     ;返回m_zxpt点的左右两个端点
    (setq m_angle (m_halfangle m_ptnearst m_zxpt)) ;返回角度
    (if (and (> m_angle (angtof "90")) (<= m_angle (angtof "270")))
      (progn    ;在二、三象限
(setq m_alignment "ml")
(setq m_rotate (+ m_angle (angtof "180")))
(setq m_wzang (- 0 (angtof "90"))) ;正90度
      )
      (progn    ;在一、四象限
(setq m_alignment "mr")
(setq m_rotate m_angle)
(setq m_wzang (angtof "90")) ;正90度
      )
    )
    (m_bzzh1_1 (list m_zxpt m_angle m_zxlength)
        (list m_wzang m_alignment m_rotate)
        m_vlaobj
        m_startlength
        m_data
    )
    (setq i (1+ i))
  )
  (setq m_zxpt (vlax-curve-getendpoint m_vlaobj))
  (setq m_ptnearst (m_searchpt m_zxpt m_vlaobj m_ptlist))
     ;返回m_zxpt点的左右两个端点
  (setq m_angle (m_halfangle m_ptnearst m_zxpt)) ;返回角度
  (if (and (> m_angle (angtof "90")) (<= m_angle (angtof "270")))
    (progn    ;在二、三象限
      (setq m_alignment "ml")
      (setq m_rotate (+ m_angle (angtof "180")))
      (setq m_wzang (- 0 (angtof "90"))) ;正90度
    )
    (progn    ;在一、四象限
      (setq m_alignment "mr")
      (setq m_rotate m_angle)
      (setq m_wzang (angtof "90")) ;正90度
    )
  )
  (m_bzzh1_1 (list m_zxpt m_angle m_zxlength)
      (list m_wzang m_alignment m_rotate)
      m_vlaobj
      m_startlength
      m_data
  )
)

(defun m_bzzh1_halfangle_Selectpoint (tm_zxlength  m_startlength
          m_ptlist    m_vlaobj
          m_data    /
          m_pt    m_zxpt
          m_ptnearst   m_wzang
          m_angle    m_alignment
          m_rotate
         )
  ;;桩号0＋000.000的点的文字对齐点的转角增量m_wzang
  (while (setq m_pt (getpoint "\n桩号点<退出>:"))
    (setq m_zxpt (vlax-curve-getclosestpointto m_vlaobj m_pt))
    (if (> (distance m_pt m_zxpt) 0.0)
      (princ "所选点不在曲线上!")
      (progn
(setq m_ptnearst (m_searchpt m_zxpt m_vlaobj m_ptlist))
     ;返回两个端点
(setq m_angle (m_halfangle m_ptnearst m_zxpt)) ;返回角度
(if (and (> m_angle (angtof "90")) (<= m_angle (angtof "270")))
     ;在二、三象限
   (progn   ;在二、三象限
     (setq m_alignment "ml")
     (setq m_rotate (+ m_angle (angtof "180")))
     (setq m_wzang (- 0 (angtof "90"))) ;正90度
   )
   (progn   ;在一、四象限
     (setq m_alignment "mr")
     (setq m_rotate m_angle)
     (setq m_wzang (angtof "90")) ;正90度
   )
)
;;m_zhlist—>(要标注的桩线点m_zxpt 桩线旋转角m_zxrotate 桩线长度m_zxlength)
;;m_wzlist—>(文字对齐点的与桩线角度的增量m_wzang 文字对齐方式m_wzalignment 文字旋转角m_wzrotate)
(m_bzzh1_1 (list m_zxpt m_angle m_zxlength)
     (list m_wzang m_alignment m_rotate)
     m_vlaobj
     m_startlength
     m_data
)
      )
    )
  )
)
(defun m_bzzh1_halfangle_Allvertex (tm_zxlength  m_startlength
        m_ptlist  m_vlaobj
        m_data  /
        n   m_zxpt
        m_ptnearst  m_angle
        m_wzang  m_alignment
        m_rotate
       )
     ;桩号0＋000.000的点的文字对齐点的转角增量m_wzang
  (setq n 0)
  (repeat (length m_ptlist)
    (setq m_zxpt (car (nth n m_ptlist)))
    (if (> (abs (- m_startlength
     (vlax-curve-getdistatpoint m_vlaobj m_zxpt)
  )
    )
    0.00000001
)    ;不是起点
      (progn
(setq m_ptnearst (m_searchpt m_zxpt m_vlaobj m_ptlist))
     ;返回两个端点
(setq m_angle (m_halfangle m_ptnearst m_zxpt)) ;返回角度
(if (and (> m_angle (angtof "90")) (<= m_angle (angtof "270")))
     ;在二、三象限
   (progn   ;在二、三象限
     (setq m_alignment "ml")
     (setq m_rotate (+ m_angle (angtof "180")))
     (setq m_wzang (- 0 (angtof "90"))) ;正90度
   )
   (progn   ;在一、四象限
     (setq m_alignment "mr")
     (setq m_rotate m_angle)
     (setq m_wzang (angtof "90")) ;正90度
   )
)
;;m_zhlist—>(要标注的桩线点m_zxpt 桩线旋转角m_zxrotate 桩线长度m_zxlength)
;;m_wzlist—>(文字对齐点的与桩线角度的增量m_wzang 文字对齐方式m_wzalignment 文字旋转角m_wzrotate)
(m_bzzh1_1 (list m_zxpt m_angle m_zxlength)
     (list m_wzang m_alignment m_rotate)
     m_vlaobj
     m_startlength
     m_data
)
      )
    )
    (setq n (1+ n))
  )
)
(defun m_halfangle (m_ptlist   m_pt   /      m_angle
      m_ang1     m_ang2   m_angtmp   m_angtmp1
      m_angtmp2  m_pt1   m_pt2      m_ptcenter
     )
  ;;根据给定点表求半角,点表由有两个或三个子点组成的表组成
  (if (= 2 (length m_ptlist))  ;只有2个点
    (if (= 2 (length (car m_ptlist))) ;圆弧段
      (setq m_angle (angle (cadar m_ptlist) m_pt))
      (progn
(setq m_angle (+ (angle (car (nth 0 m_ptlist))
    (car (nth 1 m_ptlist))
    )
    (/ pi 2.0)
        )
)
(if (> m_angle (* 2 pi))
   (setq m_angle (* (* 2 pi)
      (- (/ m_angle (* 2 pi))
         (fix (/ m_angle (* 2 pi)))
      )
   )
   )
)
(if (< m_angle 0)
   (setq m_angle (+ (* 2 pi) m_angle))
)
      )
    )
    (progn    ;有3个点
      (if (= 2 (length (nth 0 m_ptlist))) ;圆弧段
(progn
   (setq m_ptcenter (cadr (nth 0 m_ptlist)))
   (setq m_angtmp1 (angle (car (nth 1 m_ptlist))
     (cadr (nth 0 m_ptlist))
     )
   )
   (setq m_angtmp2 (angle (car (nth 1 m_ptlist))
     (car (nth 0 m_ptlist))
     )
   )
   (if (> (abs (- m_angtmp1 m_angtmp2)) (angtof "180"))
     (setq m_angtmp (angtof "270"))
     (setq m_angtmp (angtof "90"))
   )
   (if (> m_angtmp1 m_angtmp2)
     (setq m_ang1 (- (angle (car (nth 1 m_ptlist))
       (cadr (nth 0 m_ptlist))
       )
       m_angtmp
    )
     )
     (setq m_ang1 (+ (angle (car (nth 1 m_ptlist))
       (cadr (nth 0 m_ptlist))
       )
       m_angtmp
    )
     )
   )
)
(setq
   m_ang1 (angle (car (nth 1 m_ptlist)) (car (nth 0 m_ptlist)))
)    ;直线段
      )
      (if (= 2 (length (nth 1 m_ptlist))) ;圆弧段
(progn
   (setq m_ptcenter (cadr (nth 1 m_ptlist)))
   (setq m_angtmp1 (angle (car (nth 1 m_ptlist))
     (cadr (nth 1 m_ptlist))
     )
   )
   (setq m_angtmp2 (angle (car (nth 1 m_ptlist))
     (car (nth 2 m_ptlist))
     )
   )
   (if (> (abs (- m_angtmp1 m_angtmp2)) (angtof "180"))
     (setq m_angtmp (angtof "270"))
     (setq m_angtmp (angtof "90"))
   )
   (if (> m_angtmp1 m_angtmp2)
     (setq m_ang2 (- (angle (car (nth 1 m_ptlist))
       (cadr (nth 1 m_ptlist))
       )
       m_angtmp
    )
     )
     (setq m_ang2 (+ (angle (car (nth 1 m_ptlist))
       (cadr (nth 1 m_ptlist))
       )
       m_angtmp
    )
     )
   )
)
(setq
   m_ang2 (angle (car (nth 1 m_ptlist)) (car (nth 2 m_ptlist)))
)    ;直线段
      )
      (if (> m_ang1 (* 2 pi))
(setq
   m_ang1 (* (* 2 pi)
      (- (/ m_ang1 (* 2 pi)) (fix (/ m_ang1 (* 2 pi))))
   )
)
      )
      (if (> m_ang2 (* 2 pi))
(setq
   m_ang2 (* (* 2 pi)
      (- (/ m_ang2 (* 2 pi)) (fix (/ m_ang2 (* 2 pi))))
   )
)
      )
      (if (< m_ang1 0)
(setq m_ang1 (+ (* 2 pi) m_ang1))
      )
      (if (< m_ang2 0)
(setq m_ang2 (+ (* 2 pi) m_ang2))
      )
      (setq m_angle (/ (+ m_ang1 m_ang2) 2.0))
      (if (or (= 2 (length (nth 0 m_ptlist)))
       (= 2 (length (nth 1 m_ptlist)))
   )
(progn    ;有圆弧段
   (setq m_pt1 (polar m_pt m_angle 10))
   (setq m_pt2 (polar m_pt (+ (angtof "180") m_angle) 10))
   (if
     (< (distance m_pt1 m_ptcenter) (distance m_pt2 m_ptcenter))
      (setq m_angle (+ (angtof "180") m_angle))
   )
)
(progn    ;均为直线段
   (setq m_pt1 (car (nth 0 m_ptlist)))
   (setq m_pt2 (car (nth 2 m_ptlist)))
   (setq m_ptcenter (polar m_pt m_angle 10))
   (if
     (< (+ (abs (- (angle m_pt m_pt1) (angle m_pt m_ptcenter)))
    (abs (- (angle m_pt m_pt2) (angle m_pt m_ptcenter)))
        )
        (angtof "180")
     )
      (setq m_angle (+ (angtof "180") m_angle))
   )
)
      )
    )
  )
  m_angle
)
(defun m_searchpt (m_pt m_vlaobj m_ptlist / m_len n)
  ;;根据给定曲线上任意一点，返回此点左右两个顶点,如果给定点为多义线的顶点，则返回连续3个顶点
  (setq m_len (vlax-curve-getdistatpoint m_vlaobj m_pt))
  (setq n 0)
  (while (> m_len
     (vlax-curve-getdistatpoint m_vlaobj (car (nth n m_ptlist)))
  )
    (setq n (1+ n))
  )
  (if (= m_len
  (vlax-curve-getdistatpoint m_vlaobj (car (nth n m_ptlist)))
      )     ;点为多义线的顶点
    (if (= 0 n)    ;第一个顶点
      (list (nth 0 m_ptlist) (nth 1 m_ptlist)) ;第一个顶点
      (if (= (1+ n) (length m_ptlist)) ;最后一个顶点
(list (nth (1- n) m_ptlist) (nth n m_ptlist)) ;最后一个顶点
(list (nth (1- n) m_ptlist)
       (nth n m_ptlist)
       (nth (1+ n) m_ptlist)
)    ;不是最后一个顶点
      )
    )
    (list (nth (1- n) m_ptlist) (nth n m_ptlist)) ;点不是多义线的顶点
  )
)
(defun m_searchhd1 (m_plent   /  m_pttab   m_pt1     m_pt2
      m_tmp     m_ptlist m_xc   m_radius  m_pt3
      m_ptcenter
     )
  ;;返回多段线的顶点表，返回(...(点座标)...(点座标 圆心点座标)...)表中如果有两个点，则表示有圆弧段,第二个点为圆心
  (setq m_pttab (entget m_plent))
  (while (setq m_pt1 (assoc '10 m_pttab))
    (setq m_tmp (assoc '42 m_pttab))
    (if (/= 0.0 (cdr m_tmp))
      (if (setq m_pt2 (assoc '10 (cdr (member m_pt1 m_pttab)))) ;下一点
(progn
   (setq m_xc (distance (cdr m_pt1) (cdr m_pt2))) ;弦长
   (setq
     m_radius (abs (/ (* m_xc (1+ (* (cdr m_tmp) (cdr m_tmp))))
        (* 4 (cdr m_tmp))
     )
       )
   )    ;半径R
   (setq m_pt3 (polar (cdr m_pt1)
        (angle (cdr m_pt1) (cdr m_pt2))
        (/ m_xc 2.0)
        )
   )
   (if (> 0.0 (cdr m_tmp))
     (setq m_ptcenter
     (polar m_pt3
     (- (angle (cdr m_pt1) (cdr m_pt2))
        (angtof "90")
     )
     (- m_radius
        (/ (* (abs (cdr m_tmp)) m_xc) 2.0)
     )
     )
     )
     (setq m_ptcenter
     (polar m_pt3
     (- (angle (cdr m_pt1) (cdr m_pt2))
        (angtof "270")
     )
     (- m_radius
        (/ (* (abs (cdr m_tmp)) m_xc) 2.0)
     )
     )
     )
   )
   (setq m_ptlist (append m_ptlist
     (list (list (cdr m_pt1) m_ptcenter))
    )
   )
)
(setq m_ptlist (append m_ptlist (list (list (cdr m_pt1)))))
      )
      (setq m_ptlist (append m_ptlist (list (list (cdr m_pt1)))))
    )
    (setq m_pttab (cdr (member m_tmp m_pttab)))
  )
  m_ptlist
)
(princ"\n桩号标注：ZH")


(defun m_DrawText(m_FontStyle m_AlignmentStyle m_AlignmentPosition m_TextHeight m_Rotate m_Text / ztb) 
  ;;根据给定的字体、对齐方式、对齐点、字高、旋转角绘制文字 
  (if (setq ztb (tblsearch "style" m_FontStyle));;如果指定字体m_zt存在 
    (if (> (cdr (assoc '40 ztb)) 0.0);;如果指定字体的默认高度大于0.0 
      (progn 
        (command "text" "s" m_FontStyle m_AlignmentStyle m_AlignmentPosition m_Rotate m_Text);;先按默认高度写字 
        (setq ztb (entget(entlast))) 
        (setq ztb (subst (cons 40 m_TextHeight) (assoc '40 ztb) ztb));;改变默认高度为指定字高m_zg 
        (entmod ztb) 
      ) 
      (command "text" "s" m_FontStyle m_AlignmentStyle m_AlignmentPosition m_TextHeight m_Rotate m_Text);;如果指定字体默认高度等于0.0则按指定字高直接写字 
    ) 
    (m_DrawText "STANDARD" m_AlignmentStyle m_AlignmentPosition m_TextHeight m_Rotate m_Text);;如果指定字体不存在，则用标准字体STANDARD写字 
  )         
) 

(defun m_DrawText (m_FontStyle m_AlignmentStyle m_AlignmentPosition m_TextHeight m_Rotate m_Text / ztb)
  (if (setq ztb (tblsearch "style" m_FontStyle))
    (if (> (cdr (assoc '40 ztb)) 0.0)
      (progn
        (command "text" "s" m_FontStyle m_AlignmentStyle m_AlignmentPosition m_Rotate m_Text )
        (setq ztb (entget (entlast)))
        (setq ztb (subst (cons 40 m_TextHeight) (assoc '40 ztb) ztb))
        (entmod ztb)
      )
      (command "text" "s" m_FontStyle m_AlignmentStyle m_AlignmentPosition m_TextHeight m_Rotate m_Text )
    )
    (m_DrawText "STANDARD" m_AlignmentStyle m_AlignmentPosition m_TextHeight m_Rotate m_Text )
  )
)
