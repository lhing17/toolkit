(ns clj.brainfuck)

(defn add [n]
  (cond (nil? n) 1
        (= 255 n) 0
        :else (inc n)))

(defn sub [n]
  (if (nil? n) 255 (dec n)))

(defn- ord->str [n] (str (char n)))

(defn- nil-or-zero? [n] (or (nil? n) (zero? n)))

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
       (take idx)
       (filter identity)
       last
       inc
       ))

(defn execute-string
  "Evaluate the Brainfuck source code in `source` using `input` as a source of
  characters for the `,` input command.

  Either returns a sequence of output characters, or `nil` if there was
  insufficient input."
  [source input]
  (loop [data {} result "" index 0 src-index 0 in input]
    ;(print data result index (nth source src-index) in "\n")
    (if (= (count source) src-index)
      result
      (case (nth source src-index)
        \> (recur data result (inc index) (inc src-index) in)
        \< (recur data result (dec index) (inc src-index) in)
        \+ (recur (update data index add) result index (inc src-index) in)
        \- (recur (update data index sub) result index (inc src-index) in)
        \. (recur data (str result (ord->str (get data index 0))) index (inc src-index) in)
        \, (recur (assoc data index (int (first in))) result index (inc src-index) (next in))
        \[ (if (zero? (get data index 0))
             (recur data result index (next-rb source src-index) in)
             (recur data result index (inc src-index) in))
        \] (if (not= 0 (get data index 0))
             (recur data result index (last-lb source src-index) in)
             (recur data result index (inc src-index) in))
        )))
  )

(defn -main [& args]
  (let [in (-> (range 1 256) vec (conj 0)
               (->> (map char) (apply str)))]
    (println in)
    (print (= in (execute-string "+[,.]" in))))
  )