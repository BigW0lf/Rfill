;;; =========================================================
;;; CONTOUR-PIECE.lsp  |  Commande : CP  v4
;;; Contour de pièce — épaisseur cloisons détectée auto
;;; Calque cloisons par défaut : "Cloisons"
;;; =========================================================

;;;--- Géométrie 2D ----------------------------------------

(defun _cp:cross (a b)
  (- (* (car a)(cadr b)) (* (cadr a)(car b))))

(defun _cp:dot (a b)
  (+ (* (car a)(car b)) (* (cadr a)(cadr b))))

(defun _cp:norm (v / k)
  (setq k (sqrt (_cp:dot v v)))
  (if (> k 1e-10) (list (/ (car v) k) (/ (cadr v) k)) '(1.0 0.0)))

;; Normale droite = vers l'extérieur d'un polygone CCW
(defun _cp:right-n (v) (list (cadr v) (- (car v))))

;; Intersection de deux droites P1+t*D1 et P2+s*D2
(defun _cp:isect (p1 d1 p2 d2 / dp den tt)
  (setq dp  (list (- (car p2)(car p1))(- (cadr p2)(cadr p1)))
        den (_cp:cross d1 d2))
  (if (< (abs den) 1e-10)
    (list (/ (+ (car p1)(car p2)) 2.0)(/ (+ (cadr p1)(cadr p2)) 2.0))
    (progn
      (setq tt (/ (_cp:cross dp d2) den))
      (list (+ (car p1)(* tt (car d1)))
            (+ (cadr p1)(* tt (cadr d1)))))))

;; Aire signée (> 0 = CCW)
(defun _cp:area (pts / n s i p1 p2)
  (setq n (length pts) s 0.0 i 0)
  (while (< i n)
    (setq p1 (nth i pts) p2 (nth (rem (1+ i) n) pts))
    (setq s (+ s (_cp:cross p1 p2)))
    (setq i (1+ i)))
  (* s 0.5))

;; Distance perpendiculaire point → droite infinie (a,b)
(defun _cp:dpl (p a b / dx dy len)
  (setq dx (- (car b)(car a))
        dy (- (cadr b)(cadr a))
        len (sqrt (+ (* dx dx)(* dy dy))))
  (if (< len 1e-6) 1e10
    (abs (/ (- (* dx (- (cadr p)(cadr a)))
               (* dy (- (car p)(car a)))) len))))

;; Parallèle si dot des normalisés > 0.97 (~14°)
(defun _cp:par-p (d1 d2)
  (> (abs (_cp:dot (_cp:norm d1) (_cp:norm d2))) 0.97))

;; Projection scalaire de p sur segment [a,b] entre 0 et 1
(defun _cp:proj (p a b / dx dy len2 t)
  (setq dx (- (car b)(car a)) dy (- (cadr b)(cadr a))
        len2 (+ (* dx dx)(* dy dy)))
  (if (< len2 1e-10) 0.5
    (/ (+ (* (- (car p)(car a)) dx)(* (- (cadr p)(cadr a)) dy)) len2)))

;;;--- Entités AutoCAD -------------------------------------

(defun _cp:verts (ent / r)
  (setq r '())
  (foreach x (entget ent)
    (if (= (car x) 10)
      (setq r (append r (list (list (cadr x)(caddr x)))))))
  r)

(defun _cp:segs (ent / ed tp pts n i cl segs)
  (setq ed (entget ent) tp (cdr (assoc 0 ed)) segs '() pts '())
  (cond
    ((= tp "LINE")
     (list (list (list (cadr (assoc 10 ed))(caddr (assoc 10 ed)))
                 (list (cadr (assoc 11 ed))(caddr (assoc 11 ed))))))
    ((= tp "LWPOLYLINE")
     (foreach x ed
       (if (= (car x) 10)
         (setq pts (append pts (list (list (cadr x)(caddr x)))))))
     (setq n (length pts) cl (= (logand (cdr (assoc 70 ed)) 1) 1) i 0)
     (while (< i (if cl n (1- n)))
       (setq segs (append segs (list (list (nth i pts)(nth (rem (1+ i) n) pts)))))
       (setq i (1+ i)))
     segs)
    (t '())))

(defun _cp:mkpoly (pts layer / n data i p)
  (setq n (length pts)
        data (list '(0 . "LWPOLYLINE") '(100 . "AcDbEntity")
                   (cons 8 layer) '(100 . "AcDbPolyline")
                   (cons 90 n) '(70 . 1) '(43 . 0.0)))
  (setq i 0)
  (while (< i n)
    (setq p (nth i pts))
    (setq data (append data (list (cons 10 (list (car p)(cadr p))))))
    (setq i (1+ i)))
  (entmake data))

(defun _cp:mklayer (name)
  (if (not (tblsearch "LAYER" name))
    (entmake (list '(0 . "LAYER") '(100 . "AcDbSymbolTableRecord")
                   '(100 . "AcDbLayerTableRecord") (cons 2 name)
                   '(70 . 0) '(62 . 3) '(6 . "Continuous")))))

;;;--- Précharge tous les segments du calque cloisons ------
;;; Retourne une liste plate de (pa pb) pour chaque segment trouvé

(defun _cp:load-cloison-segs (lc / ss len i ent et result)
  (setq result '())
  (setq ss (ssget "_X" (list (cons 8 lc)
                              (cons 0 "LINE,LWPOLYLINE"))))
  (if (not ss)
    (prompt (strcat "\n[DBG] Aucune entite LINE/LWPOLYLINE sur calque [" lc "] dans le dessin !"))
    (progn
      (setq len (sslength ss) i 0)
      (prompt (strcat "\n[DBG] " (itoa len) " entite(s) trouvee(s) sur calque [" lc "]"))
      (while (< i len)
        (setq ent (ssname ss i))
        (foreach seg (_cp:segs ent)
          (setq result (append result (list seg))))
        (setq i (1+ i)))
      (prompt (strcat "\n[DBG] " (itoa (length result)) " segment(s) charges"))))
  result)

;;;--- Calcule l'offset pour un segment du contour ---------
;;;
;;;  Cherche parmi tous les segments cloisons lesquels sont :
;;;    1. Parallèles au segment contour
;;;    2. Proches (dist perp < maxDist depuis le milieu)
;;;  Parmi ceux-là : sépare coincidents (dist < tol) et partenaires
;;;  Retourne ep/2

;; Projection du point p sur le segment [a,b], retourne t (0=a, 1=b)
(defun _cp:proj-t (p a b / dx dy len2)
  (setq dx (- (car b)(car a)) dy (- (cadr b)(cadr a))
        len2 (+ (* dx dx)(* dy dy)))
  (if (< len2 1e-10) 0.5
    (/ (+ (* (- (car p)(car a)) dx)(* (- (cadr p)(cadr a)) dy)) len2)))

(defun _cp:offset (v1 v2 csegs / mx my sdir maxDist tolC minD coincident result ea eb sd dd pt)
  (setq mx        (/ (+ (car v1)(car v2)) 2.0)
        my        (/ (+ (cadr v1)(cadr v2)) 2.0)
        sdir      (_cp:norm (list (- (car v2)(car v1))(- (cadr v2)(cadr v1))))
        maxDist   2.0   ; épaisseur max cherchée
        tolC      0.01  ; dist < tolC = coïncident
        coincident nil
        minD      1e10
        result    0.0)

  (foreach seg csegs
    (setq ea (car seg) eb (cadr seg)
          sd (list (- (car eb)(car ea))(- (cadr eb)(cadr ea))))
    (setq dd (_cp:dpl (list mx my) ea eb))
    ;; Condition : parallèle + proche + milieu du contour EN FACE de ce segment
    (if (and (_cp:par-p sdir sd)
             (< dd maxDist)
             ;; Le milieu du segment contour doit se projeter sur le candidat
             (setq pt (_cp:proj-t (list mx my) ea eb))
             (> pt -0.3)(< pt 1.3))
      (cond
        ((< dd tolC) (setq coincident t))
        ((< dd minD) (setq minD dd)))))

  (cond
    ((and coincident (< minD 1e9))
     (setq result (/ minD 2.0))
     (prompt (strcat "\n[DBG] seg(" (rtos mx 2 2) "," (rtos my 2 2)
                     ") CLOISON ep=" (rtos minD 2 4)
                     " offset=" (rtos result 2 4))))
    (coincident
     (prompt (strcat "\n[DBG] seg(" (rtos mx 2 2) "," (rtos my 2 2) ") ligne seule, offset=0")))
    (t
     (prompt (strcat "\n[DBG] seg(" (rtos mx 2 2) "," (rtos my 2 2) ") pas de cloison"))))
  result)

;;;--- Commande principale CP ------------------------------

(defun c:CP (/ lc lo pt ent verts n i v1 v2 d nrm off op offsets olines nverts prev cur np csegs aire)

  (initget 0)
  (setq lc (getstring T "\nCalque des cloisons <Cloisons> : "))
  (if (= lc "") (setq lc "Cloisons"))

  (initget 0)
  (setq lo (getstring T "\nCalque de sortie <CONTOURS> : "))
  (if (= lo "") (setq lo "CONTOURS"))

  ;; Précharger tous les segments cloisons UNE SEULE FOIS
  (setq csegs (_cp:load-cloison-segs lc))
  (if (null csegs)
    (progn (prompt "\nAucun segment cloison trouvé, abandon.")(exit)))

  (_cp:mklayer lo)

  ;; Boucle : continuer jusqu'à Echap ou Entrée sans clic
  (while
    (progn
      (setq pt (getpoint "\nCliquer dans une pièce [Entrée pour terminer] : "))
      pt)

    (command "_-BOUNDARY" pt "")
    (setq ent (entlast))
    (if (not (= (cdr (assoc 0 (entget ent))) "LWPOLYLINE"))
      (prompt "\nErreur : zone non fermée, clic suivant.")
      (progn
        (setq verts (_cp:verts ent) n (length verts))
        (entdel ent)
        (if (< n 3)
          (prompt "\nContour invalide, clic suivant.")
          (progn
            (setq aire (_cp:area verts))
            (prompt (strcat "\n[DBG] Contour : " (itoa n) " sommets, aire=" (rtos aire 2 3)))
            (if (< aire 0.0)(setq verts (reverse verts)))

            (setq offsets '() i 0)
            (while (< i n)
              (setq v1 (nth i verts) v2 (nth (rem (1+ i) n) verts))
              (setq offsets (append offsets (list (_cp:offset v1 v2 csegs))))
              (setq i (1+ i)))

            (setq olines '() i 0)
            (while (< i n)
              (setq v1  (nth i verts)
                    v2  (nth (rem (1+ i) n) verts)
                    d   (list (- (car v2)(car v1))(- (cadr v2)(cadr v1)))
                    nrm (_cp:right-n (_cp:norm d))
                    off (nth i offsets)
                    op  (list (+ (car v1)(* off (car nrm)))
                              (+ (cadr v1)(* off (cadr nrm)))))
              (setq olines (append olines (list (list op d))))
              (setq i (1+ i)))

            (setq nverts '() i 0)
            (while (< i n)
              (setq prev (nth (rem (+ n (1- i)) n) olines)
                    cur  (nth i olines)
                    np   (_cp:isect (car prev)(cadr prev)(car cur)(cadr cur)))
              (setq nverts (append nverts (list np)))
              (setq i (1+ i)))

            (_cp:mkpoly nverts lo)
            (prompt (strcat "\nContour trace sur [" lo "] — clic suivant ou Entree pour terminer.")))))))

  (prompt "\nTermine.")
  (princ))

;;;--- Commande CPINFO : diagnostic calques ----------------

(defun c:CPINFO (/ lc ss len i ed tp)
  (initget 0)
  (setq lc (getstring T "\nCalque a inspecter <Cloisons> : "))
  (if (= lc "") (setq lc "Cloisons"))
  (setq ss (ssget "_X" (list (cons 8 lc))))
  (if (not ss)
    (prompt (strcat "\n=> AUCUNE entite sur [" lc "] dans tout le dessin"))
    (progn
      (setq len (sslength ss))
      (prompt (strcat "\n=> " (itoa len) " entite(s) sur [" lc "]"))
      (setq i 0)
      (while (< i (min len 5))
        (setq ed (entget (ssname ss i)))
        (prompt (strcat "\n   " (cdr (assoc 0 ed))
                        " @ (" (rtos (cadr (assoc 10 ed)) 2 2)
                        "," (rtos (caddr (assoc 10 ed)) 2 2) ")"))
        (setq i (1+ i)))))
  (setq ss (ssget "_X" '((0 . "INSERT"))))
  (if ss
    (progn
      (prompt (strcat "\n=> " (itoa (sslength ss)) " INSERT(s)/XREF(s)"))
      (setq i 0)
      (while (< i (min (sslength ss) 5))
        (setq ed (entget (ssname ss i)))
        (prompt (strcat "\n   bloc=[" (cdr (assoc 2 ed)) "] calque=[" (cdr (assoc 8 ed)) "]"))
        (setq i (1+ i)))))
  (princ))

(prompt "\nCONTOUR-PIECE.lsp chargé  —  commandes : CP  et  CPINFO")
(princ)
