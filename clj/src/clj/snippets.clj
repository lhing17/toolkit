(ns clj.snippets
  (:require [clojure.string :as str]))

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