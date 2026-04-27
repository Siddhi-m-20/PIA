package Mathematics;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Random;
import java.util.Scanner;
import java.util.Vector;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Assign11_03040503_102 {
	public static Scanner sc = new Scanner(System.in);
	public static Random random = new Random();

	public static void main(String args[]) throws IOException, FileNotFoundException {
		String filename = "C:\\Siddhi\\Maths_Excelsheet\\VML_03040503_102_Assign11_Siddhi.xlsx";
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
			int chr = random.nextInt(cht.length);
			int chr11 = random.nextInt(cht.length);
			int chr1 = random.nextInt(cht1.length);
			int chr2 = random.nextInt(cht1.length);
			int chr3 = random.nextInt(cht2.length);

			if (chr11 != chr && chr1 != chr2 && checkIfNoCommonFactor(a, b)) {
				String ch = cht[chr];
				String ch11 = cht[chr11];
				String ch1 = cht1[chr1];
				String ch2 = cht1[chr2];
				String ch3 = cht2[chr3];

				String Question = "Which of the following expressions is not a linear polynomial in one variable?<br>#खाली दिलेल्या बैजिक राशींमधून अशी राशी शोधा जी एका चलातील रेषीय बहुपदी नाही.<br> ";
				String Correct_ans = "";
				int corr = random.nextInt(4);
				switch (corr) {
				case 0:
					Correct_ans = "$" + a + ch + "+" + b + "-" + c + ch11 + "$";
					break;
				case 1:
					Correct_ans = "$" + d + ch1 + "-" + c + ch2 + "-" + e + "$";
					break;
				case 2:
					Correct_ans = "$" + b + ch1 + "^2+" + f + ch2 + "+" + g + "$";
					break;
				case 3:
					Correct_ans = "$" + h + ch3 + "+" + d + ch3 + "^2-" + k + ch3 + "^3$";
					break;
				}
				String wrong_ans1 = "$" + a + ch + "+" + b + "$<br>";
				String wrong_ans2 = "$" + c + "+" + d + ch11 + "$<br>";
				String wrong_ans3 = "$" + e + ch1 + "-" + f + "$<br>";
				String wrong_ans4 = "$" + g + "-" + h + ch2 + "$<br>";
				String wrong_ans5 = "None of the given<br>#दिलेल्या पैकी कोणतीही नाही.<br> ";
				String wrong_ans6 = "All of the given<br>#दिलेल्या पैकी सर्व.<br>";
				String sol1 = "Ans : " + Correct_ans +"<br>"
						+ "For any algebraic expression to be a polynomial in one variable, it must be <br>$i)$ in one variable only and<br>$ii)$ the power of the variable must be $1$<br>Accordingly all the given expressions satisfy these conditions, either directly or after simplification.<br>But for "
						+ Correct_ans + ", this expression does not satisfy the required conditions.<br>";

				String sol2 = "";
				switch (corr) {
				case 0:
					sol2 = "As " + Correct_ans + " is an expression having $2$ variables as, $" + ch + "$ and $" + ch11
							+ "$ respectively.<br>";
					break;
				case 1:
					sol2 = "As " + Correct_ans + " is an expression having $2$ variables as, $" + ch1 + "$ and $" + ch2
							+ "$ respectively.<br>";
					break;
				case 2:
					sol2 = "As " + Correct_ans + " is an expression having $2$ variables as, $" + ch1 + "$ and $" + ch2
							+ "$ respectively." + " <br>And, the power of variable is $2$ due to the term $" + b + ch1 + "^2"
							+ "$.<br>";
					break;
				case 3:
					sol2 = "As " + Correct_ans
							+ " is though an expression in one variable, but the power of variable is $2$ and $3$ due to the terms $"
							+ d + ch3 + "^2$  and  $" + "-" + k + ch3 + "^3$ respectively.<br>";
					break;
				}
				String sol3 = "Hence the answer is " + Correct_ans
						+ " is not a linear polynomial in one variable.<br>#";
				String soleng = sol1 + sol2 + sol3;
				String sola = "उत्तर : " + Correct_ans + "<br>"
						+ "कोणतीही बैजिक राशी ही एक चलातील रेषीय बहुपदी असण्यासाठी, ती राशी<br> $i)$ फक्त एक चलातीलच असायला हवी आणि <br>"
						+ "$ii)$ त्यातील चलाचा घातांक हा $1$ असाच असायला हवा.<br>"
						+ "या नुसार दिलेल्या सर्व राशी या अटी पूर्ण करतात किंवा सरळ रूप दिल्यावर पूर्ण करतात.<br> परंतु "
						+ Correct_ans + " ही राशी दिलेल्या अटी पूर्ण करत नाही. <br>";
				String solb = "";
				switch (corr) {
				case 0:
					solb = "कारण " + Correct_ans + " या राशीत दोन चल $" + ch + "$ आणि $" + ch11 + "$ आहेत. <br>";
					break;
				case 1:
					solb = "कारण " + Correct_ans + " या राशीत दोन चल $" + ch1 + "$ आणि $" + ch2 + "$ आहेत. <br>";
					break;
				case 2:
					solb = "कारण " + Correct_ans + " या राशीत दोन चल $" + ch1 + "$ आणि $" + ch2 + "$ आहेत. <br>"
							+ "  आणि त्या राशीतील $" + b + ch1 + "^2$ या पदाचा घातांक $2$ आहे.  <br>";
					break;
				case 3:
					solb = "कारण " + Correct_ans + " या राशीत एक चल आहे. <br>" + "तरी त्या राशीतील $"+ d + ch3 + "^2$ आणि $" + "-" + k + ch3 + "^3$ या पदात चलाचा घातांक अनुक्रमे $2$ आणि $3$ आहे. <br>";
					break;
				}
				String solc = "म्हणून " + Correct_ans + " ही एक चलातील रेषीय नसलेली राशी आहे, हे उत्तर.<br>";
				String solmar = sola + solb + solc;
				String Solution = soleng + solmar;
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
//			row.createCell(14).setCellValue(" ");
				row.createCell(15).setCellValue("siddhimanjarekar274@gmail.com");
				row.createCell(16).setCellValue(Solution);
//			row.createCell(17).setCellValue(" ");
				row.createCell(18).setCellValue(102);

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
