package Mathematics;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashSet;
import java.util.Random;
import java.util.Scanner;
import java.util.Vector;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Assign16_03040503_106 {
	public static Scanner sc = new Scanner(System.in);
	public static Random random = new Random();

	public static void main(String args[]) throws IOException, FileNotFoundException {
		String filename = "C:\\Siddhi\\Maths_Excelsheet\\VML_03040503_106_Assign16_Siddhi.xlsx";
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Instruction");

		XSSFSheet sheet1 = workbook.createSheet("Questions");

		// Adding header to the Questions sheet
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

		int q = sc.nextInt();
		int i = 1;
		do {
			Random r = new Random();
			int min = 2;
			int max = 20;
			int min1 = 11;
			int max1 = 30;
			int a = random.nextInt(max - min + 1) + min;
			int b = random.nextInt(max - min + 1) + min;
			int c = random.nextInt(max - min + 1) + min;
			int d = random.nextInt(max - min + 1) + min;
			int e = random.nextInt(max - min + 1) + min;
			int f = random.nextInt(max1 - min1 + 1) + min1;
			int g = random.nextInt(max1 - min1 + 1) + min1;
			int h = random.nextInt(max1 - min1 + 1) + min1;
			int k = random.nextInt(max1 - min1 + 1) + min1;

			String cht[] = { "u", "v", "w" };
			String cht1[] = { "x", "y", "z" };
			String cht2[] = { "a", "b", "c" };
			String cht3[] = { "p", "q", "r", "s", "t" };
			int chr = random.nextInt(cht.length);
			int chr11 = random.nextInt(cht.length);
			int chr1 = random.nextInt(cht1.length);
			int chr2 = random.nextInt(cht1.length);
			int chr3 = random.nextInt(cht2.length);

			int chr4 = random.nextInt(cht3.length);
			int chr5 = random.nextInt(cht3.length);
			int chr6 = random.nextInt(cht3.length);

			if (chr11 != chr && chr1 != chr2 && checkIfNoCommonFactor(a, b)) {
				String ch = cht[chr];
				String ch11 = cht[chr11];
				String ch1 = cht1[chr1];
				String ch2 = cht1[chr2];
				String ch3 = cht2[chr3];
				String ch4 = cht3[chr4];
				String ch5 = cht3[chr5];
				String ch6 = cht3[chr6];
				String Question = "Which of the following expressions is not a cubic polynomial in one variable?<br>#खाली दिलेल्या बैजिक राशींमधून अशी राशी शोधा जी एका चलातील घन बहुपदी नाही.<br> ";
				String Correct_ans = "";
				int corr = random.nextInt(6);
				switch (corr) {
				case 0:
					Correct_ans = "$" + a + ch + "^3 +" + b + ch11 + "^3$";
					break;
				case 1:
					Correct_ans = "$" + a + ch1 + "^2 -" + b + ch2 + "+" + c + "$";
					break;
				case 2:
					Correct_ans = "$" + a + ch2 + "-" + b + ch1 + "$";
					break;
				case 3:
					Correct_ans = "$" + a + ch2 + "^2 -" + b + ch1 + "-" + c + "$";
					break;
				case 4:
					Correct_ans = "$" + a + ch4 + " -" + ch5 + "^2$";
					break;
				case 5:
					Correct_ans = "$" + ch3 + "^2 +" + c + "$";
					break;
				case 6:
					Correct_ans = "$" + a + ch4 + "-" + b + ch5 + "+" + c + ch6 + "$";
					break;
				case 7:
					Correct_ans = "$" + a + ch + "+" + b + "$";
					break;
				}
				String wrong_ans1 = "$" + ch + "^3 +" + d + ch + "^2 +" + e + ch + "+" + f + "$<br>";
				String wrong_ans2 = "$" + g + ch11 + "^3$<br>";
				String wrong_ans3 = "$" + ch1 + "^3  -" + c + ch1 + "+" + d + "$<br>";
				String wrong_ans4 = "$" + c + ch2 + "^3 +" + b + ch2 + "^2 -" + a + ch2 + "$<br>";
				String wrong_ans5 = "$" + ch3 + "^3 -" + a + ch3 + "^2 -" + k + ch3 + "+" + d + "$<br>";
				String wrong_ans6 = "$" + g + ch + "^3 " + "-" + h + "$<br>";
				String Solution = "Ans: " + Correct_ans
						+ "<br>For any expression to be a cubic polynomial in one variable, it must satisfy following three conditions simultaneously. <br>$i)$ It must be in one variable only.<br>$ii)$ The highest power of the variable in the expression must be three. <br> $iii)$ The coefficient of the term with the power $3$ must not be zero. <br>From the given options only "
						+ Correct_ans
						+ " does not  satisfy all the three conditions simultaneously, mentioned above.<br> All remaining options do satisfy the given conditions simultaneously.<br> Hence the correct answer is "
						+ Correct_ans + "<br>#उत्तर : " + Correct_ans
						+ "<br>कोणतीही बैजिक राशी एका चलातील घन राशी असण्यासाठी, खालील तीन अटींचे एकाच वेळी पालन होणे आवश्यक आहे. <br>$i)$ यात एकच चल असावा. <br>$ii)$ या राशी मधील चलाचा मोठ्यात मोठा घातांक $3$ असावा. <br>$iii)$ $3$ घातांक असलेल्या पदातील सहगुणक कधीही शून्य नसावा. <br>या नुसार दिलेल्या पर्यायांपैकी फक्त "
						+ Correct_ans + " ही राशी दिलेल्या तीनही अटींचे एकाच वेळी पालन करत नाही.<br>म्हणून "
						+ Correct_ans + " हे बरोबर उत्तर आहे. <br>";
				Vector<String> vec = new Vector<String>();
				vec.add(wrong_ans1);
				vec.add(wrong_ans2);
				vec.add(wrong_ans3);
				vec.add(wrong_ans4);
				vec.add(wrong_ans5);
				vec.add(wrong_ans6);
				int wrg = r.nextInt(6);
				HashSet<String> set = new HashSet<String>();
				for (int j = 0; j < vec.size(); j++) {
					set.add(vec.elementAt(j));
				}
				XSSFRow row = sheet1.createRow(i);
				row.createCell(0).setCellValue(i);
				row.createCell(1).setCellValue("Text");
				row.createCell(2).setCellValue(1);
				row.createCell(3).setCellValue("03040503");
				row.createCell(4).setCellValue(Question);
				row.createCell(5).setCellValue(Correct_ans + "<br>");
				row.createCell(9).setCellValue(vec.elementAt(wrg));
				vec.removeElementAt(wrg);
				wrg = r.nextInt(5);
				row.createCell(10).setCellValue(vec.elementAt(wrg));
				vec.removeElementAt(wrg);
				wrg = r.nextInt(4);
				row.createCell(11).setCellValue(vec.elementAt(wrg));
				vec.removeElementAt(wrg);
				row.createCell(12).setCellValue(90);
				row.createCell(13).setCellValue(2);
//		row.createCell(14).setCellValue(" ");
				row.createCell(15).setCellValue("siddhimanjarekar274@gmail.com");
				row.createCell(16).setCellValue(Solution);
//		row.createCell(17).setCellValue(" ");
				row.createCell(18).setCellValue(106);

				if (Correct_ans == wrong_ans1 || Correct_ans == wrong_ans2 || Correct_ans == wrong_ans3
						|| Correct_ans == wrong_ans4 || Correct_ans == wrong_ans5 || Correct_ans == wrong_ans6
						|| wrong_ans1 == wrong_ans2 || wrong_ans1 == wrong_ans3 || wrong_ans1 == wrong_ans4
						|| wrong_ans1 == wrong_ans5 || wrong_ans1 == wrong_ans6 || wrong_ans2 == wrong_ans3
						|| wrong_ans2 == wrong_ans4 || wrong_ans2 == wrong_ans5 || wrong_ans2 == wrong_ans6
						|| wrong_ans3 == wrong_ans4 || wrong_ans3 == wrong_ans4 || wrong_ans3 == wrong_ans6
						|| wrong_ans4 == wrong_ans5 || wrong_ans4 == wrong_ans6 || wrong_ans5 == wrong_ans6) {
					System.out.println("duplicate" + i);
					i--;
				}
			} else {
				continue;
			}
			i++;
		} while (i < q + 1);

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

	public static boolean checkIfNoCommonFactor(int a, int b) {
		return gcd(a, b) == 1;
	}

	private static int gcd(int a, int b) {
		while (b != 0) {
			int temp = b;
			b = a % b;
			a = temp;
		}
		return a;
	}
}
