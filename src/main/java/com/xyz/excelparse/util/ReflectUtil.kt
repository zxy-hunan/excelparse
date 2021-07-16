package com.xyz.excelparse.util

import android.util.Log
import java.lang.reflect.Method


class ReflectUtil {
//ss
    companion object {

        fun <T> listToBean(list: List<String>, clazz: Class<T>): T {
            val obj = clazz.newInstance()
            val declaredFields = clazz.declaredFields
            declaredFields?.let {
                var i = 0
                Log.e("test","declaredFields=${declaredFields?.size}")
                while (i < declaredFields.size) {
                    if (list.size != declaredFields.size) {
                        break
                    } else {
                        val value = list[i]?.toString()
                        var method: Method? = null
                        when (declaredFields[i].genericType.toString()) {
                            "class java.lang.String"  -> method = clazz.getMethod(
                                "set" + getMethodName(declaredFields[i].name),
                                String::class.java
                            )
                            "class java.lang.Integer" -> method = clazz.getMethod(
                                "set" + getMethodName(declaredFields[i].name),
                                Int::class.java
                            )
                            "class java.lang.Boolean" -> method = clazz.getMethod(
                                "set" + getMethodName(declaredFields[i].name),
                                Boolean::class.java
                            )
                        }
                        method?.invoke(obj, value)
                        i++
                    }
                }
            }
            return obj as T
        }


        private fun getMethodName(fieldName: String): String? {
            val bytes = fieldName.toByteArray()
            bytes[0] = (bytes[0].minus('a'.toByte()) + 'A'.toByte()).toByte()
            return String(bytes)
        }

    }


}