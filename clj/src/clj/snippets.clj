(ns clj.snippets
  (:require [clojure.string :as str])
  (:import (java.util Properties)))

(defn underline->camel [name]
  "下划线转驼峰"
  ;; str/replace的第二个参数为正则表达式时，最后一个参数可以传一个函数，该函数的参数为正则表达式捕获的分组
  (str/replace name #"_(\w)" #(str/upper-case (% 1))))


(defn first-index-matching [pred coll]
  "第一个符合条件的元素的下标"
  (ffirst
    (drop-while #((complement pred) (second %))
                (map-indexed vector coll))))

(defn first-index-matching [pred coll]
  "第一个符合条件的元素的下标，另一种实现"
  (->> coll
       (map-indexed #(when (pred %2) %))
       (filter identity)                                    ;; 去掉nil
       (first)))

(defn revrot [strng sz]
  "反转或翻转"
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

(defn- main [& args]
  (print (revrot "733049910872815764", 5))
  )

(main)