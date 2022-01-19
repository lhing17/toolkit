1. 在REPL中加载某个文件

```clojure
(use 'XXX.XXX :reload-all)
```

2. defprotocol用于声明协议，extend-protocol用于为不同类型扩展协议。也可以直接在defrecord中扩展协议。

3. 在ns（命名空间）中，引入clojure模块和Java模块的方式分别为：

```clojure
(:require [clojure.java.io :as io]
  [ring.util.codec :refer [assoc-conj]])
(:import [java.io PipedInputStream PipedOutputStream ByteArrayInputStream File Closeable IOException])

```

其中:as为别名，:refer为引用某个/某些方法

4. defn-用于声明私有函数

5. 声明重载函数示例：

```clojure
 (defn string-input-stream
       "Returns a ByteArrayInputStream for the given String."
       {:added "1.1"}
       ([^String s]
        (ByteArrayInputStream. (.getBytes s)))
       ([^String s ^String encoding]
        (ByteArrayInputStream. (.getBytes s encoding))))
```

6. -> ->> doto as->的区别

1) ->串联一组函数，将每一个函数的返回值作为下一个函数的第一个参数

```clojure
   (-> (.lastModified file)
       (/ 1000) (long) (* 1000)
       (java.util.Date.))
```

2) ->> 串联一组函数，将每一个函数的返回值作为下一个函数的最后一个参数
3) doto后第一个参数将作为后续所有函数的第一个参数及返回值
4) as-> 将第一个参数用第二个参数的符号表示，后续该符号变为每个函数的返回值

```clojure
 (as-> 5 $
       (* 4 $)
       (- 14 $)
       (/ $ 2))
```

7. try-catch-finally示例：

```clojure
(try
  (.close ^java.io.Closeable stream)
  (catch IOException _ nil)
  (finally (.close stream))) 
```   

8. proxy用于实现java的一个接口，类似于java中的匿名内部类，示例：

```clojure
   (defn- ^AbstractHandler proxy-handler [handler]
          (proxy [AbstractHandler] []
                 (handle [_ ^Request base-request ^HttpServletRequest request response]
                         (when-not (= (.getDispatcherType request) DispatcherType/ERROR)
                                   (let [request-map (servlet/build-request-map request)
                                         response-map (handler request-map)]
                                        (servlet/update-servlet-response response response-map)
                                        (.setHandled base-request true))))))
```

9. 利用reduce和assoc将request中所有参数加入到map中

```clojure
(defn- get-headers
       "Creates a name/value map of all the request headers."
       [^HttpServletRequest request]
       (reduce
         (fn [headers, ^String name]
             (assoc headers
                    (.toLowerCase name Locale/ENGLISH)
                    (->> (.getHeaders request name)
                         (enumeration-seq)
                         (string/join ","))))
         {}
         (enumeration-seq (.getHeaderNames request))))
```

10. 解构map中的值：

```clojure
 let [{:keys [status headers body]} response-map]
```

函数式编程和数据建模已经给了我们很强的表达力，并且使得我们可以把代码里面的重复模式抽象出来。宏应该被当做我们的终极武器，用它来简化控制流，添加一些语法糖以消除重复代码以及一些难看的代码。

11. 使用file-seq来实现文件夹的遍历。
12. 使用(seq coll)来判断一个集合非空，不要用(not (empty? coll))
13. 当不想让结果为惰性集合时，使用mapv、filterv这类的函数代替map和filter。
14. 将字符序列转成字符串：(apply str coll)。
15. 将二元vector的序列转成map: (into {} coll)。
16. 按数量快速分组使用partition，如(partition 5 coll)即为5个一组。注意partition会舍弃不足5个值的组，如果希望保留，应该用partition-all。
17. 可以将中间结果写在临时文件中，这样可以清晰地看到代码执行的过程。
18. 如果需要一个函数，不管参数是什么，返回结果都一样，用constantly: (constantly x)，返回值是一个函数。
19. 如果需要一个函数，返回参数本身，用identity: (identity x)，返回值是x。identity有时可以用来判断布尔结果，因为在clojure中，只有nil和false在布尔判断中为否，所以，可以用类似(filter identity coll)的方式来去掉coll中的nil和false值。
20. 将map后的结果连接到一块用mapcat，(mapcat f coll)相当于(apply concat (map f coll))。
21. 取多元序列中的数据可以用get-in: (get-in m ks)，更新数据可以用update-in: (update-in m ks f)。
22. defmulti和defmethod配合使用，可以定义多重方法，用于处理不同类型的参数。如：

```clojure
 ;this example illustrates that the dispatch type
;does not have to be a symbol, but can be anything (in this case, it's a string)

(defmulti greeting
          (fn [x] (get x "language")))

;params is not used, so we could have used [_]
(defmethod greeting "English" [params]
           "Hello!")

(defmethod greeting "French" [params]
           "Bonjour!")

;;default handling
(defmethod greeting :default [params]
           (throw (IllegalArgumentException.
                    (str "I don't know the " (get params "language") " language"))))

;then can use this like this:
(def english-map {"id" "1", "language" "English"})
(def french-map {"id" "2", "language" "French"})
(def spanish-map {"id" "3", "language" "Spanish"})

=> (greeting english-map)
"Hello!"
=> (greeting french-map)
"Bonjour!"
=> (greeting spanish-map)
java.lang.IllegalArgumentException: I don't know the Spanish language

```

23. 常用变量命名 rdr -> reader hdr -> header c -> char coll -> collection
24. 对于互相引用的函数，可以在上方使用declare先进行声明，如：

```clojure
;; Ahead decl because some fns call into each other.
(declare parse parse-input parse-file tag-content)
```

25. 可以用for形式来实现笛卡尔积，如
```clojure

(for [x "abc" y "123"] (str x y)) ;;("a1" "a2" "a3" "b1" "b2" "b3" "c1" "c2" "c3")

```

26. 可以灵活使用reduce-kv，将map中的key-value对进行约减。
