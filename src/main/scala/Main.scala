import java.io.{FileInputStream, FileOutputStream}
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import scala.util.Random

object Main {
  def main(args: Array[String]) {
    println("start")

    if(args.length != 2){
      println("args error")
      return
    }

    // e.g.
    // /Users/tsukaby/IdeaProjects/poi-scala-sample/src/main/resources/template.xlsx
    val tempalteFilePath = args{0}
    // /Users/tsukaby/IdeaProjects/poi-scala-sample/result.xlsx
    val outputFilePath = args{1}

    val is = new FileInputStream(args{0})
    val wb = new XSSFWorkbook(is)

    val sheet = wb.getSheet("Sheet1")

    val r = new Random

    for(i <- 1 until 11){
      val row = sheet.getRow(i)

      val cell = row.createCell(0)
      cell.setCellValue("Student" + i)

      for(j <- 1 until 4){
        val cell = row.createCell(j)
        val point = r.nextInt(100)
        cell.setCellValue(point)
      }
    }

    val os = new FileOutputStream({args{1}})
    wb.write(os)

    println("end")
  }
}
