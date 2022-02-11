(ns clj.snippets
  (:require [clojure.string :as str])
  (:import (java.util Properties)))

(defn underline->camel
  "下划线转驼峰"
  [name]
  ;; str/replace的第二个参数为正则表达式时，最后一个参数可以传一个函数，该函数的参数为正则表达式捕获的分组
  (str/replace name #"_(\w)" #(str/upper-case (% 1))))


(defn first-index-matching
  "第一个符合条件的元素的下标"
  [pred coll]
  (ffirst
   (drop-while #((complement pred) (second %))
               (map-indexed vector coll))))

(defn first-index-matching
  "第一个符合条件的元素的下标，另一种实现"
  [pred coll]
  (->> coll
       (map-indexed #(when (pred %2) %))
       (filter identity) ;; 去掉nil
       (first)))

(defn revrot
  "反转或翻转"
  [strng sz]
        (letfn [(cubic [n] (* n n n))
          (rev-or-rot [coll]
            (let [sum (apply + (map (comp cubic read-string str) coll))]
              (if (even? sum) (reverse coll) (concat (rest coll) [(first coll)]))))]
    (if (or (<= sz 0)
            (> sz (count strng))
            (empty? strng))
      ""
      (->> strng
           (partition sz)
           (map (comp #(apply str %) rev-or-rot))
           (apply str)))))


(defn- ^Properties as-properties
  "将map转成java.util.Properties实例。
  Convert any seq of pairs to a `java.util.Properties` instance."
  [m]
  (let [p (Properties.)]
    (doseq [[k v] m]
      (.setProperty p (name k) (str v)))
    p))


;; 以下为A*寻路算法相关的函数
(defn neighbors
  "寻找近邻，返回合法的近邻点
  `size` -- 矩阵的边长
  `yx`   -- 要寻找近邻的点"
  ([size yx]
   (neighbors [[-1 0] [1 0] [0 -1] [0 1]] size yx))
  ([deltas size yx]
   (filter (fn [new-yx]
             (every? #(< -1 % size) new-yx))
           (map #(mapv + yx %) deltas))))

(defn path-cost
  "计算目前已遍历的路径成本
  `node-cost`      -- 现有的路径成本
  `cheapest-nbr`   -- 成本最小的近邻，格式为{:cost 1, ...}
  "
  [node-cost cheapest-nbr]
  (+ node-cost (or (:cost cheapest-nbr) 0)))

(defn estimate-cost
  "估计余下路径的成本函数
  `step-cost-est`  -- 估算每步的成本
  `size`           -- 矩阵的边长
  `y`              -- y坐标
  `x`              -- x坐标
  "
  [step-cost-est size y x]
  (* step-cost-est
     (- (+ size size) y x 2)))

(defn total-cost
  "计算路径估计成本
  `newcost`        -- 新成本
  `step-cost-est`  -- 估算每步的成本
  `size`           -- 矩阵的边长
  `y`              -- y坐标
  `x`              -- x坐标
  "
  [newcost step-cost-est size y x]
  (+ newcost (estimate-cost step-cost-est size y x)))

(defn min-by
  "取出最小值的函数
  `f`              -- 预处理函数
  `coll`           -- 要处理的集合
  "
  [f coll]
  (when (seq coll)
    (reduce (fn [min other]
              (if (> (f min) (f other))
                other
                min))
            coll)))

(defn astar
  "A*寻路算法"
  [start-yx step-est cell-costs]
  (let [size (count cell-costs)]
    (loop [steps 0
           routes (vec (repeat size (vec (repeat size nil))))
           work-todo (sorted-set [0 start-yx])]
      (print steps "->" routes work-todo "\n")
      (if (empty? work-todo)
        ; 检查是否完成
        [(peek (peek routes)) :steps steps]
        (let [[_ yx :as work-item] (first work-todo)
              rest-work-todo (disj work-todo work-item)
              nbr-yxs (neighbors size yx)
              cheapest-nbr (min-by :cost (keep #(get-in routes %) nbr-yxs))
              newcost (path-cost (get-in cell-costs yx) cheapest-nbr)
              oldcost (:cost (get-in routes yx))]
          (if (and oldcost (>= newcost oldcost))
            (recur (inc steps) routes rest-work-todo)
            (recur (inc steps)
                   (assoc-in routes yx {:cost newcost :yxs (conj (:yxs cheapest-nbr []) yx)})
                   (into rest-work-todo (map (fn [w]
                                               (let [[y x] w]
                                                 [(total-cost newcost step-est size y x) w]))
                                             nbr-yxs)))))))))

(def world [[1 1 1 1 1]
            [999 999 999 999 1] 
            [1 1 1 1 1]
            [1 999 999 999 999]
            [1 1 1 1 1]])

(defn- next-rb [src idx]
  (->> src
       (map-indexed #(when (= \] %2) %))
       (drop (inc idx))
       (filter identity)
       first
       inc)) 

(defn- last-lb [src idx]
  (->> src
       (map-indexed #(when (= \[ %2) %))
       (take (inc idx))
       (filter identity)
       last
       inc
       ))

;; Is there a core fn to change from [{:name "Foo" :value 1}, {:name "Bar" :value 2}] to {"Foo": 1, "Bar": 2}?
(defn convert [coll]
  (into {} (map (juxt :name :value) coll)))

(comment
  (def coll [{:name "Foo" :value 1}, {:name "Bar" :value 2}])
  (convert coll) ;; => {"Foo" 1, "Bar" 2}
)


(defn- -main [& args]
  ;(print (revrot "733049910872815764", 5))
  ;(println (astar [0 0] 900 world))
  (print (last-lb "][3[[[[[4234235]343]" 5)))

  

