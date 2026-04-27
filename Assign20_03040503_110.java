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

public class Assign20_03040503_110 {
	public static Scanner sc = new Scanner(System.in);
	public static Random random = new Random();

	public static void main(String args[]) throws IOException, FileNotFoundException {
		String filename = "C:\\Siddhi\\Maths_Excelsheet\\VLab_03040503_110_Assign20_Siddhi.xlsx";
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
			int chr4 = random.nextInt(cht1.length);
			if (chr11 != chr && chr1 != chr2 && checkIfNoCommonFactor(a, b)) {
				String ch = cht[chr];
				String ch11 = cht[chr11];
				String ch1 = cht1[chr1];
				String ch2 = cht1[chr2];
				String ch3 = cht2[chr3];
				String ch4 = cht1[chr4];

				String Question = "Identify from the given set of polynomials in one variable as cubic polynomials.<br>#खाली दिलेल्या बैजिक राशींपैकी कोण कोणत्या राशी एका चलातील घन बहुपदी आहेत?<br>";
				String Correct_ans1 = "";
				String Correct_ans2 = "";
				int corr = random.nextInt(4);
				switch (corr) {
				case 0:
					Correct_ans1 = "$" + ch + "+" + e + ch + "^3 +" + f + ch + "+" + g + "$";
					Correct_ans2 = "$" + a + ch11 + "^3 -" + b + ch11 + "^2 -" + d + "$";
					break;
				case 1:
					Correct_ans1 = "$" + ch1 + " -" + c + ch1 + "^3 +" + d + "$";
					Correct_ans2 = "$" + a + ch2 + "+" + b + ch2 + "^2 -" + e + ch2 + "^3$";
					break;
				case 2:
					Correct_ans1 = "$" + ch3 + "^3 -" + b + ch3 + "^2 -" + c + ch3 + "+" + d + "$";
					Correct_ans2 = "$" + a + ch + "^3 " + "-" + e + "$";
					break;
				case 3:
					Correct_ans1 = "$" + a + ch3 + "^3$";
					Correct_ans2 = "$" + d + ch2 + "^3 -" + e + "$";
					break;
				}
				String wrong_ans1 = "$" + b + ch1 + "^2 -" + ch1 + "+" + g + "$<br>";
				String wrong_ans2 = "$" + a + ch + "^3 +" + b + ch11 + "^3$<br>";
				String wrong_ans3 = "$" + ch2 + ch1 + "^2 -" + e + ch1 + ch2 + "^2$<br>";
				String wrong_ans4 = "$" + k + ch3 + "+" + g + "$<br>";
				String wrong_ans5 = "$" + c + ch2 + "-" + h + ch1 + "+" + k + ch4 + "$<br>";
				String wrong_ans6 = "$" + h + ch2 + "^4 -" + g + ch1 + "$<br>";

				String Solution = "Ans $1$ :  " + Correct_ans1 + " <br>Ans $2$ :  " + Correct_ans2
						+ " <br>According to the definition, for any algebraic expression to be a cubic polynomial in one variable, it must be in one variable only and the highest power of the variable must be $3$. <br> Accordingly, we can see that, as per definition "
						+ Correct_ans1 + " and " + Correct_ans2
						+ "   are the polynomials which satisfy the given conditions. <br>Both of them are in $1$ variable and the highest power of the variable is $3$.<br>"
						+ "Of the remaining, either they are not in one variable or the degree of the polynomial is not $3$.<br>$\\therefore$ Answer $1$ :  "
						+ Correct_ans1 + ".<br>Answer $2$ :  " + Correct_ans2 + ".<br>#उत्तर  $1:$   " + Correct_ans1
						+ "<br>उत्तर  $2:$  " + Correct_ans2
						+ " <br>कोणतीही बैजिक राशी एका चलातील घन बहुपदी असण्यासाठी, तिच्यात एकच चल असायला हवा आणि यातील चलाचा सगळ्यात मोठा घातांक $3$ असायला हवा. <br> या नुसार आपल्याला हे लक्षात येते की, "
						+ Correct_ans1 + "  आणि " + Correct_ans2
						+ " या दोनच राशी एका चलातील घन बहुपदी आहेत.<br>कारण या दोन्ही राशी एकाच चलातील आहेत आणि त्यातील चलाचा मोठ्यात मोठा घातांक $3$ आहे.<br>उरलेल्या राशी पैकी, एकतर त्यांचा घातांक $3$ नाही किंवा त्यात एकापेक्षा जास्त चल आहेत.<br>$\\therefore$   उत्तर  $1: $"
						+ Correct_ans1 + "  आहे<br>  उत्तर  $2 :$ " + Correct_ans2 + " आहे.<br> ";
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
				row.createCell(2).setCellValue(5);
				row.createCell(3).setCellValue("03040503");
				row.createCell(4).setCellValue(Question);
				row.createCell(5).setCellValue(Correct_ans1 + "<br>");
				row.createCell(6).setCellValue(Correct_ans2 + "<br>");
				row.createCell(9).setCellValue(vec.elementAt(wrg));
				vec.removeElementAt(wrg);
				wrg = r.nextInt(5);
				row.createCell(10).setCellValue(vec.elementAt(wrg));
				vec.removeElementAt(wrg);
				/*
				 * wrg = r.nextInt(4); row.createCell(11).setCellValue(vec.elementAt(wrg));
				 * vec.removeElementAt(wrg);
				 */
				row.createCell(12).setCellValue(90);
				row.createCell(13).setCellValue(3);
//			row.createCell(14).setCellValue(" ");
				row.createCell(15).setCellValue("siddhimanjarekar274@gmail.com");
				row.createCell(16).setCellValue(Solution);
//			row.createCell(17).setCellValue(" ");
				row.createCell(18).setCellValue(110);

				if (Correct_ans1 == wrong_ans1 || Correct_ans1 == wrong_ans2 || Correct_ans1 == wrong_ans3
						|| Correct_ans1 == wrong_ans4 || Correct_ans1 == wrong_ans5 || Correct_ans1 == wrong_ans6
						|| Correct_ans2 == wrong_ans1 || Correct_ans2 == wrong_ans2 || Correct_ans2 == wrong_ans3
						|| Correct_ans2 == wrong_ans4 || Correct_ans2 == wrong_ans5 || Correct_ans2 == wrong_ans6
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
