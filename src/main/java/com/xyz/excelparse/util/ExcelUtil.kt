package com.xyz.excelparse.util

import android.os.Handler
import android.util.Log
import org.apache.poi.POIXMLDocument
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.openxml4j.opc.OPCPackage
import org.apache.poi.poifs.filesystem.POIFSFileSystem
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.FormulaEvaluator
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileInputStream
import java.io.InputStream
import java.io.PushbackInputStream


class ExcelUtil {
}

inline fun <reified T> readExcel(file: File?, inputsteam: InputStream?, mHandler: Handler): MutableList<T>? {
    var carInfoList: MutableList<T>? = null
    when (file ?: inputsteam) {
        is InputStream -> {
            inputsteam?.let {
                carInfoList = excuteSteam1(inputsteam)
                mHandler.sendEmptyMessage(0)
            }
        }

        is File -> {
            file?.let {
                val inputsteam = FileInputStream(it)
                carInfoList = excuteSteam(inputsteam)
            }
        }

        null -> println()
    }
    return carInfoList
}


fun getBook(steam: InputStream): Workbook? {
    var tempSteam: InputStream? = null
    var xssfWorkbook: Workbook? = null
    steam?.let {
        if (!steam.markSupported()) {
            tempSteam = PushbackInputStream(steam, 8)
        }
        if (POIFSFileSystem.hasPOIFSHeader(steam)) {
            xssfWorkbook = HSSFWorkbook(steam)
        } else if (POIXMLDocument.hasOOXMLHeader(steam)) {
            xssfWorkbook = XSSFWorkbook(OPCPackage.open(steam))
        } else {
        }
    }
    return xssfWorkbook
}


fun <T> excuteSteam(steam: InputStream): MutableList<T> {
    val carInfoList = mutableListOf<T>()
    val xssfWorkbook: Workbook? = getBook(steam)
    steam?.let {
        val sheet = xssfWorkbook?.getSheetAt(0)
        val formulaEvaluator: FormulaEvaluator? =
            xssfWorkbook?.creationHelper?.createFormulaEvaluator()
        sheet?.let { it ->

            for (row in it) {
                val count = row.physicalNumberOfCells
                var i = 0
                while (i < count) {
                    val cell = row.getCell(i)
                    val cellValue = formulaEvaluator?.evaluate(cell)
                    println("cellvalue:${cellValue?.stringValue.toString()}")

                    i++
                }
//                carInfoList.add(null)
            }
        }
    }
    return carInfoList
}

fun readExcel(inputsteam: InputStream) {
    inputsteam?.let {
        val xssfWorkbook = XSSFWorkbook(inputsteam)
        val sheet = xssfWorkbook.getSheetAt(0)
        val formulaEvaluator: FormulaEvaluator =
            xssfWorkbook.getCreationHelper().createFormulaEvaluator()
        sheet?.let {
            for (row in it) {
                val count = row.physicalNumberOfCells
                var i = 0
                while (i < count) {
                    val cell = row.getCell(i)
                    val cellValue = formulaEvaluator.evaluate(cell)
                    println("cellvalue:${cellValue.toString()}")
                    i++
                }
            }
        }
    }
}


inline fun <reified T> excuteSteam1(steam: InputStream): MutableList<T> {
    val rowList = mutableListOf<T>()
    val xssfWorkbook: Workbook? = getBook(steam)
    steam?.let {
        val sheet = xssfWorkbook?.getSheetAt(0)
        val formulaEvaluator: FormulaEvaluator? =
            xssfWorkbook?.creationHelper?.createFormulaEvaluator()
        sheet?.let { it ->
            var obj: T? = null
            for (row in it) {
                val count = row.physicalNumberOfCells
                var i = 0
                val rowColumnList = mutableListOf<String>()
                while (i < count) {
                    val cell = row.getCell(i)
                    val cellValue = formulaEvaluator?.evaluate(cell)
                    println("cellvalue:${cellValue?.stringValue.toString()}")
                    rowColumnList.apply {
                        cellValue?.let { it ->
                            var value: String? = null
                            when (it.cellType) {
                                Cell.CELL_TYPE_BOOLEAN -> value = it.booleanValue?.toString()
                                Cell.CELL_TYPE_NUMERIC -> value = it.numberValue?.toString()
                                Cell.CELL_TYPE_STRING -> value = it.stringValue?.toString()
                                null -> value = ""
                            }
                            value?.let { it ->
                                add(it)
                            }
                        }
                    }
                    i++
                }
                Log.e("test", "rowColumnList=${rowColumnList?.size}")
                obj = ReflectUtil.listToBean(rowColumnList, T::class.java)
                obj?.let {
                    rowList.add(obj)
                }
            }
        }
    }
    return rowList
}

