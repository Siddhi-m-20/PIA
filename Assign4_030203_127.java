package Mathematics;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Random;
import java.util.Scanner;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Assign4_030203_127 {
	public static Scanner scanner = new Scanner(System.in);
	public static Random random = new Random();

	public static void main(String[] args) throws IOException, FileNotFoundException {
		String filename = "C:\\Siddhi\\Maths_Excelsheet\\VML_0302030_127_Assign4_Siddhi.xlsx";
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
		for (int i = 0; i < header.length; i++) {
			rowhead.createCell(i).setCellValue(header[i]);
		}

		// Taking input for the number of questions to generate
		System.out.println("How many questions do you want to enter:");
		int numQuestions = scanner.nextInt();

		int mapsize, mapsizeafter;
		HashMap<String, Integer> map = new HashMap<String, Integer>();
		int i = 1;
		do {
			String[] girlenglish = { "Arya", "Swati", "Radha", "Prachi", "Aditi", "Akshara", "Kavyaa", "Swara",
					"Sharayu", "sayali" };
			String[] girlmarathi = { "आर्या", "स्वाती", "राधा", "प्राची", "अदिती", "अक्षरा", "काव्या", "स्वरा", "शरयू",
					"सायली" };
			String boysenglish[] = { "Abhinav", "Harsh", "Adarsh", "Shaym", "Rajveer", "Amit", "Rajvardhan", "Ashok",
					"Omkar", "Kunal" };
			String boysmarathi[] = { "अभिनव", "हर्ष", "आदर्श", "श्याम", "राजवीर", "अमित", "राजवर्धन", "अशोक", "ओंकार",
					"कुणाल" };
			int p = random.nextInt(girlenglish.length);
			String name = " ";
			String name1 = " ";
			String gender = " ";
			String gender1 = " ";
			String gender2 = " ";
			String gender3 = " ";
			String gender4 = " ";
			String gender5 = "";
			int g = random.nextInt(2);
			if (g == 0) {
				name = girlenglish[p];
				name1 = girlmarathi[p];
				gender = "she";
				gender3 = "her";
				gender1 = "तिचे";
				gender2 = "तिच्या";
				gender4 = "वयाची";
				gender5 = "Her";
			} else {
				name = boysenglish[p];
				name1 = boysmarathi[p];
				gender = "he";
				gender3 = "his";
				gender1 = "त्याचे";
				gender2 = "त्याच्या";
				gender4 = "वयाचा";
				gender5 = "His";
			}
			// Generating random values
			int minA = 1, maxA = 40, minB = 2, maxB = 9, minC = 1, maxC = 20;
			int a = random.nextInt(maxA - minA + 1) + minA;
			int b = random.nextInt(maxB - minB + 1) + minB;
			int c = random.nextInt(maxC - minC + 1) + minC;
			int mul = b * c;
			int coex = b - 1;
			int mulplusa = mul + a;
			int answer = mulplusa / coex;
			if (mulplusa % coex == 0) {

				String questionText = "After " + year(a) + " , " + name + " shall be $" + b + "$ times as old as "
						+ gender + " was " + year(c) + " ago. Find " + gender3 + " present age.<br >#$" + a
						+ "$ वर्षानंतर  " + name1 + "  " + gender2 + " $" + c + "$ वर्षांपूर्वीच्या वयाच्या $" + b
						+ "$ पट  " + gender4 + " होईल. तर " + gender1 + " आजचे वय काय?<br>";
				String correct_ans = "$" + answer + "$ years<br>#$" + answer + "$ वर्षे <br>";
				String wrong_ans1 = "$" + (answer - 2) + "$ years<br>#$" + (answer - 2) + "$ वर्षे <br>";
				String wrong_ans2 = "$" + (answer + 1) + "$ years<br>#$" + (answer + 1) + "$ वर्षे <br>";
				String wrong_ans3 = "$" + (answer - 1) + "$ years<br>#$" + (answer - 1) + "$ वर्षे <br>";

				// Generating solution
				String solution = "Ans : " + year(answer) + "<br>Let us assume present age of " + name
						+ " as $x$.<br>After " + year(a) + " " + gender3 + " age will be $ x + " + a + "$.<br>"
						+ gender5 + " age " + year(c) + " ago was $ x - " + c + "$.<br>But as given $" + b + "(x-" + c
						+ ") = x + " + a + "$<br>We will solve this equation for $x$<br>We get $" + b + "x - " + mul
						+ " = x + " + a + "$ . . . .  by opening the bracket.<br>$\\therefore " + b + "x - x = " + a
						+ " + " + mul + "$<br>$\\Rightarrow  " + hide1(coex) + " =" + mulplusa + "$<br>"
						+ hideans(coex, answer) + "$\\therefore$ present age of " + name + " is " + year(answer)
						+ " is the answer.<br>#उत्तर  : $" + answer + "$ वर्षे <br>" + name1
						+ "चे आजचे वय $x$ मानू.<br>$" + a + "$ वर्षांनंतर " + gender1 + " वय $x + " + a + "$ होईल.<br>$"
						+ c + "$ वर्षांपूर्वी " + gender1 + " वय $x - " + c + "$ असणार.<br>परंतु दिल्यानुसार $" + b
						+ "( x - " + c + ") = x + " + a
						+ "$<br>हे समीकरण आपण $x$ साठी सोडवले असता<br>आपल्याला कंस सोडवून $" + b + "x - " + mul
						+ " = x +" + a + "$ असे मिळते.<br>$\\therefore " + b + "x - x = " + a + " + " + mul
						+ "$<br>$\\Rightarrow  " + hide1(coex) + " = " + mulplusa + "$<br>" + hideans(coex, answer)
						+ "$\\therefore$ " + name1 + "चे आजचे वय $" + answer + "$ वर्षे आहे हे उत्तर.<br>";

				// Create row
				XSSFRow row = sheet1.createRow(i);
				row.createCell(0).setCellValue(i);
				row.createCell(1).setCellValue("Text");
				row.createCell(2).setCellValue(1);
				row.createCell(3).setCellValue("0302030");
				row.createCell(4).setCellValue(questionText);
				row.createCell(5).setCellValue(correct_ans);
				row.createCell(9).setCellValue(wrong_ans1);
				row.createCell(10).setCellValue(wrong_ans2);
				row.createCell(11).setCellValue(wrong_ans3);
				row.createCell(12).setCellValue(90);
				row.createCell(13).setCellValue(4);
				row.createCell(15).setCellValue("siddhimanjarekar274@gmail.com");
				row.createCell(16).setCellValue(solution);
				row.createCell(18).setCellValue(127);

				// Check for duplicate questions
				mapsize = map.size();
				map.put(questionText, i);
				mapsizeafter = map.size();

				// In Java, a map can consist of virtually any number of key-value pairs, but
				// the keys must always be unique — non-repeating.
				if (mapsize == mapsizeafter) {
					System.out.println("duplicate Question" + i + ". " + questionText);
					i--;
				}

				if (correct_ans == wrong_ans1 || correct_ans == wrong_ans2 || correct_ans == wrong_ans3
						|| wrong_ans1 == wrong_ans2 || wrong_ans1 == wrong_ans3 || wrong_ans2 == wrong_ans3) {
					System.out.println("duplicate" + i);
					i--;
				}
			} else {
				continue;
			}

			i++;
		} while (i <= numQuestions);

		int rowTotal = sheet1.getLastRowNum();
//              System.out.println(rowTotal);
		XSSFRow row = sheet1.createRow((short) rowTotal + 1);
		row.createCell(0).setCellValue("****");
		// Writing data to the file
		FileOutputStream fileout = new FileOutputStream(filename);
		workbook.write(fileout);
		fileout.close();

		System.out.println("File created.");

	}

	public static String hide1(int a) {
		if (a == 1) {
			return "x";
		} else {
			return a + "x";
		}
	}

	public static String hideans(int a, int answer) {
		if (a == 1) {
			return "";
		} else {
			return "$\\Rightarrow x = " + answer + "$<br>";
		}
	}

	public static String year(int a) {
		if (a == 1) {
			return "$1$ year";
		} else {
			return "$" + a + "$ years";
		}
	}

}
