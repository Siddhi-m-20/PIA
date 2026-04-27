package demo_math;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Random;
import java.util.Scanner;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Assign1_03040501_105 {
	public static Scanner sc = new Scanner(System.in);
	public static Random random = new Random();

	public static void main(String args[]) throws IOException, FileNotFoundException {
		String filename = "E:\\Demo_Workspace\\VML_03040501_105_Assign1_Siddhi.xlsx";
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Instruction");

		XSSFSheet sheet1 = workbook.createSheet("Questions");

		// Adding first row in second sheet(Questions)
		String[] header = { "Sr. No", "Question Type", "Answer Type", "Topic Number", "Question (Text Only)",
				"Correct Answer 1", "Correct Answer 2", "Correct Answer 3", "Correct Answer 4", "Wrong Answer 1",
				"Wrong Answer 2", "Wrong Answer 3", "Time in seconds", "Difficulty Level",
				"Question (Image/ Audio/ Video)", "Contributor's Registered mailId", "Solution (Text Only)",
				"Solution (Image/ Audio/ Video)", "Variation Number" };
		XSSFRow rowhead = sheet1.createRow((short) 0);

		// Set height and width to the column and row
		sheet1.setColumnWidth(4, 35 * 250);
		sheet1.setColumnWidth(16, 45 * 250);

		// Adding header to the second sheet
		for (int head = 0; head < header.length; head++) {
			rowhead.createCell(head).setCellValue(header[head]);

		}

		// Taking input for number of question you want to generate
		System.out.println("How many question you want to enter:-");

		int mapsize, mapsizeafter;
		HashMap<String, Integer> map = new HashMap<String, Integer>();
		int q = sc.nextInt();
		for (int i = 1; i < q + 1; i++) {
			// Create row
			XSSFRow row = sheet1.createRow(i);
			row.createCell(0).setCellValue(i);
			row.createCell(1).setCellValue("Text");
			row.createCell(2).setCellValue(1);
			row.createCell(3).setCellValue("03040501");

			int min = 1;
			int max = 20;
			int a = random.nextInt(max - min + 1) + min;
			int b = random.nextInt(max - min + 1) + min;

			// array of variables
			String ch[] = { "a", "b", "c", "d", "f", "g", "h", "m", "n", "p", "s", "t", "u", "v", "w", "x", "y", "z" };
			int chr = random.nextInt(ch.length);

			int p = 0;
			int q1 = 0;
			if (chr == ch.length - 1) {
				chr = chr - 2;
				p = chr + 1;
				q1 = p + 1;
			} else if (chr == 0) {
				p = 1;
				q1 = p + 1;
			} else if (chr == ch.length - 2) {
				chr = chr - 1;
				p = chr + 1;
				q1 = p + 1;
			} else {
				p = chr + 1;
				q1 = p + 1;
			}
			String chrt1 = ch[chr];
			String chrt2 = ch[p];
			int asqr = a * a;
			int bsqr = b * b;

			// Generate Correct answer
			String Correct_ans = "$\\dfrac{(" + hide1a(a,chrt1) + "-" + b + ")}{(" + hide1a(a,chrt1) + "+" + b + ")}$<br>";

			// Generate wrong options
			String wrong_ans1 = "$\\dfrac{(" + hide1a(a,chrt1) + "-" + b + ")}{(" + hide1a(a,chrt1) + "+" + b + ")^2}$<br>";
			String wrong_ans2 = "$\\dfrac{(" + hide1a(a,chrt1) + "+" + b + ")}{(" + hide1a(a,chrt1) + "-" + b + ")}$<br>";
			String wrong_ans3 = "$\\dfrac{(" + a + "-" + hide1a(b,chrt1) + ")}{(" + a + "+" + hide1a(b,chrt1) + ")}$<br>";

			String Question = "Simplify the following expression:<br> $\\dfrac{" + hide1asqr(asqr,chrt1) + "-" + bsqr + "}{("
					+ hide1a(a,chrt1) + "+" + b + ")^2}$<br>#खाली दिलेल्या राशीला सरळ रूप द्या. <br>$\\dfrac{"
					+ hide1asqr(asqr,chrt1) + "-" + bsqr + "}{(" + hide1a(a,chrt1) + "+" + b + ")^2}$<br>";

			row.createCell(4).setCellValue(Question);
			row.createCell(5).setCellValue(Correct_ans);
//                 row.createCell(6).setCellValue(" ");
//                 row.createCell(7).setCellValue(" ");
//                 row.createCell(8).setCellValue(" ");
			row.createCell(9).setCellValue(wrong_ans1);
			row.createCell(10).setCellValue(wrong_ans2);
			row.createCell(11).setCellValue(wrong_ans3);
			row.createCell(12).setCellValue(60);
			row.createCell(13).setCellValue(1);
//                  row.createCell(14).setCellValue(" ");
			row.createCell(15).setCellValue("siddhimanjarekar274@gmail.com");

			// Generate Solution
			String Solution = "Ans : " + Correct_ans + "As given $\\dfrac{" + hide1asqr(asqr,chrt1) + "-" + bsqr + "}{("
					+ hide1a(a,chrt1) + "+" + b + ")^2}$<br>"
					+ "We can see that numerator is in the form of $(x^2-y^2)$ and can be factorized as <br>$(x^2-y^2) = (x+y)(x-y)....(i)$<br>"
					+ "and denominator is a perfect square of the form $(x+y)^2$ and can be written as <br>$(x+y)^2 = (x+y)(x+y)....(ii)$<br>"
					+ "$\\therefore$ by using $(i)$ and $(ii)$ we get <br>$\\dfrac{" + hide1asqr(asqr,chrt1) + "-" + bsqr
					+ "}{(" + hide1a(a,chrt1) + "+" + b + ")^2}$<br>" + "$ = \\dfrac {(" + hide1a(a,chrt1) + ")^2-(" + b + ")^2}{("
					+ hide1a(a,chrt1) + "+" + b + ")(" + hide1a(a,chrt1) + "+" + b + ")}$<br><br>" + "$=\\dfrac{(" + hide1a(a,chrt1) + "+"
					+ b + ")(" + hide1a(a,chrt1) + "-" + b + ")}{(" + hide1a(a,chrt1) + "+" + b + ")(" + hide1a(a,chrt1) + "+" + b
					+ ")}$<br>" + "Now by cancelling common factor $(" + hide1a(a,chrt1) + "+" + b
					+ ")$ from both numerator and denominator we get<br>" + "$\\therefore\\dfrac{" + hide1asqr(asqr,chrt1)
					+ "-" + bsqr + "}{(" + hide1a(a,chrt1) + "+" + b + ")^2} = \\dfrac{(" + hide1a(a,chrt1) + "-" + b + ")}{("
					+ hide1a(a,chrt1) + "+" + b + ")}$ is the answer. <br>" + "उत्तर : " + Correct_ans + "दिल्यानुसार $\\dfrac{"
					+ hide1asqr(asqr,chrt1) + "-" + bsqr + "}{(" + hide1a(a,chrt1) + "+" + b + ")^2}$<br>"
					+ "आपण हे पाहू शकतो की दिलेल्या राशीचा अंश $(x^2-y^2)$   असा आहे, आणि त्याचे आपण खालील प्रमाणे अवयव पाडू शकतो    <br>$(x^2-y^2) = (x+y)(x-y)....(i)$<br>"
					+ "दिलेल्या राशीचा छेद  $(x+y)^2$   असा आहे, जो पूर्ण वर्ग आहे आणि त्याचे सुद्धा अवयव आपण पुढील प्रमाणे लिहू शकतो  <br>$(x+y)^2 = (x+y)(x+y)....(ii)$<br>"
					+ "$\\therefore$ आता $(i)$ आणि  $(ii)$   यांचा वापर करून आपल्याला   <br>$\\dfrac{" + hide1asqr(asqr,chrt1)
					+ "-" + bsqr + "}{(" + hide1a(a,chrt1) + "+" + b + ")^2}$<br>" + "$ = \\dfrac {(" + hide1a(a,chrt1) + ")^2-(" + b
					+ ")^2}{(" + hide1a(a,chrt1) + "+" + b + ")(" + hide1a(a,chrt1) + "+" + b + ")}$<br><br>" + "$ = \\dfrac{("
					+ hide1a(a,chrt1) + "+" + b + ")(" + hide1a(a,chrt1) + "-" + b + ")}{(" + hide1a(a,chrt1) + "+" + b + ")(" + hide1a(a,chrt1)
					+ "+" + b + ")}$  असे लिहिता येते  <br>" + "आता अंश आणि छेदातून   $(" + hide1a(a,chrt1) + "+" + b
					+ ")$   हा सामायिक अवयव रद्द करून आपल्याला <br>" + "$\\therefore\\dfrac{" + hide1asqr(asqr,chrt1) + "-"
					+ bsqr + "}{(" + hide1a(a,chrt1) + "+" + b + ")^2}=\\dfrac{(" + hide1a(a,chrt1) + "-" + b + ")}{(" + hide1a(a,chrt1)
					+ "+" + b + ")}$  असे मिळते हे उत्तर. <br>"

			;
			row.createCell(16).setCellValue(Solution);
//                  row.createCell(17).setCellValue(" ");
			row.createCell(18).setCellValue(105);

			mapsize = map.size();
			map.put(Question, i);
			mapsizeafter = map.size();

			// In Java, a map can consist of virtually any number of key-value pairs, but
			// the keys must always be unique — non-repeating.
			if (mapsize == mapsizeafter) {
				System.out.println("duplicate Question" + i + ". " + Question);
				i--;
			}

			if (Correct_ans == wrong_ans1 || Correct_ans == wrong_ans2 || Correct_ans == wrong_ans3
					|| wrong_ans1 == wrong_ans2 || wrong_ans1 == wrong_ans3 || wrong_ans2 == wrong_ans3) {
				System.out.println("duplicate" + i);
				i--;
			}

		}

		int rowTotal = sheet1.getLastRowNum();
//              System.out.println(rowTotal);
		XSSFRow row = sheet1.createRow((short) rowTotal + 1);
		row.createCell(0).setCellValue("****");

		// Writing data to the file
		FileOutputStream fileout = new FileOutputStream(filename);
		workbook.write(fileout);
		fileout.close();

		System.out.println("file created");

	}

	public static String hide1a(int a,String chr) {
		if (a != 1 && a != 0)
			return a + chr;
		else if (a == 1)
			return chr;
		else if (a == 0)
			return " ";
		return null;

	}

	public static String hide1asqr(int a,String chr) {
		if (a != 1 && a != 0)
			return a + chr+"^2";
		else if (a == 1)
			return chr+"^2";
		else if (a == 0)
			return " ";
		return null;

	}
}