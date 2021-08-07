1. 在REPL中加载某个文件

```clojure
(use 'XXX.XXX :reload-all)
```

2. defprotocol用于声明协议，extend-protocol用于为不同类型扩展协议.

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