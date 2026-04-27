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

public class Assign3_0302030_124 {
	public static Scanner scanner = new Scanner(System.in);
	public static Random random = new Random();

	public static void main(String[] args) throws IOException, FileNotFoundException {
		String filename = "C:\\Siddhi\\Maths_Excelsheet\\VML_30203_124_Assign3_Siddhi.xlsx";
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
					"Sharayu", "Sayali" };
			String[] girlmarathi = { "आर्या", "स्वाती", "राधा", "प्राची", "अदिती", "अक्षरा", "काव्या", "स्वरा", "शरयू",
					"सायली" };
			String boysenglish[] = { "Abhinav", "Harsh", "Adarsh", "Rajveer", "Amit", "Rajvardhan", "Ashok", "Omkar",
					"Kunal", "Akbar" };
			String boysmarathi[] = { "अभिनव", "हर्ष", "आदर्श", "राजवीर", "अमित", "राजवर्धन", "अशोक", "ओंकार", "कुणाल",
					"अकबर" };
			String Itemlist[] = { "jam", "pickle", "chutney", "roasted peanuts", "roasted cashew nuts",
					"flavoured yogurt", "chili powder", "turmeric powder", "peanut butter", "ginger-garlic paste",
					"curd", "butter milk", "tea", "orange juice", "handwash", "extract", " tomato ketchup",
					" hibiscus -olive oil", "coconut oil", "kokum syrup" };
			String Itemlist1[] = { "जॅम", "लोणचे", "चटणी", "भाजलेले शेंगदाणे", "भाजलेले काजू", "श्रीखंड", "मसाला",
					"हळद", "शेंगदाण्याचं लोणी", "आलं-लसूणाची पेस्ट", "दही", "ताक", "चहा", "संत्र्याचं सरबत",
					"हात धुवायची जेल", "काढा", "टोमॅटो केचप", "जास्वंदीचे तेल", "नारळाचं तेल", "कोकम सरबत" };
			String verb[] = { "बनवला", "बनवले", "बनवली", "बनवले", "बनवले", "बनवले", "बनवला", "बनवली", "बनवले", "बनवली",
					"बनवले", "बनवले", "बनवला", "बनवले", "बनवली", "बनवला", "बनवले", "बनवले", "बनवले", "बनवले" };
			String verb1[] = { "भरला", "भरले", "भरली", "भरले", "भरले", "भरले", "भरला", "भरली", "भरले", "भरली", "भरले",
					"भरले", "भरला", "भरले", "भरली", "भरला", "भरले", "भरले", "भरले", "भरले" };
			String verb2[] = { "राहीला", "राहीले ", "राहीली", "राहीले ", "राहीले ", "राहीले ", "राहीला", "राहीली",
					"राहीले ", "राहीली", "राहीले ", "राहीले ", "राहीला", "राहीले ", "राहीली", "राहीला", "राहीले ",
					"राहीले ", "राहीले ", "राहीले " };
			String verb3[] = { "होता", "होते", "होती", "होते", "होते", "होते", "होता", "होती", "होते", "होती", "होते",
					"होते", "होता", "होते", "होती", "होता", "होते", "होते", "होते", "होते" };
			String verb4[] = { " भरलेला ", "भरलेले", "भरलेली", "भरलेले", "भरलेले", "भरलेले", " भरलेला ", "भरलेली",
					"भरलेले", "भरलेली", "भरलेले", "भरलेले", " भरलेला ", "भरलेले", "भरलेली", " भरलेला ", "भरलेले",
					"भरलेले", "भरलेले", "भरलेले" };
			String verb5[] = { "उरलेला", "उरलेले", "उरलेली", "उरलेले", "उरलेले", "उरलेले", "उरलेला", "उरलेली", "उरलेले",
					"उरलेली", "उरलेले", "उरलेले", "उरलेला", "उरलेले", "उरलेली", "उरलेला", "उरलेले", "उरलेले", "उरलेले",
					"उरलेले" };
			String verb6[] = { "केलेला", "केलेले", "केलेली", "केलेले", "केलेले", "केलेले", "केलेला", "केलेली", "केलेले",
					"केलेली", "केलेले", "केलेले", "केलेला", "केलेले", "केलेली", "केलेला", "केलेले", "केलेले", "केलेले",
					"केलेले" };

			int p = random.nextInt(girlenglish.length);
			int r = random.nextInt(Itemlist.length);
			String name = " ";
			String name1 = " ";
			String name2 = Itemlist[r];
			String name3 = Itemlist1[r];
			String name4 = verb[r];
			String name5 = verb1[r];
			String name6 = verb2[r];
			String name7 = verb3[r];
			String name8 = verb4[r];
			String name9 = verb5[r];
			String name10 = verb6[r];
			String gender = " ";
			String gender1 = " ";
			String gender2 = " ";
			String gender3 = " ";
			int g = random.nextInt(2);
			if (g == 0) {
				name = girlenglish[p];
				name1 = girlmarathi[p];
				gender = "she";
				gender1 = "तिने";
				gender2 = "तिच्याकडे";
				gender3 = "her";
			} else {
				name = boysenglish[p];
				name1 = boysmarathi[p];
				gender = "he";
				gender1 = "त्याने";
				gender2 = "त्याच्याकडे";
				gender3 = "his";
			}
			String unit = " ";
			String unit2 = " ";
			if (r < 11) {
				unit = "gm";
				unit2 = "ग्रॅम";
			} else {
				unit = "ml";
				unit2 = "मिली";
			}

			// Generating random values
			int minA = 1000, maxA = 5000, minB = 10, maxB = 100, minC = 100, maxC = 1000;
			int a = random.nextInt(maxA - minA + 1) + minA;
			int b = random.nextInt(maxB - minB + 1) + minB;
			int c = random.nextInt(maxC - minC + 1) + minC;

			if ((a - c) % b == 0) {
				// Constructing the question
				String questionText = name + " prepared $" + a + "$ " + unit + " " + name2
						+ " at home and filled it in containers. After giving away $" + b + "$ of the containers to "
						+ gender3 + " friends , " + gender + " still has left with " + gender3 + " $" + c + "$ " + unit
						+ " of " + name2 + ". What was the quantity of " + name2 + " " + gender
						+ " had filled in each container? <br>#" + name1 + "ने घरी $" + a + "$ " + unit2 + " " + name3
						+ " " + name4 + " आणि बाटल्यांत " + name5 + ". त्यापैकी $" + b
						+ "$ बाटल्या आपल्या मित्र-मैत्रिणींना दिल्यावर " + gender2 + " $" + c + "$ " + unit2 + " "
						+ name3 + " शिल्लक " + name6 + ". तर " + gender1 + " प्रत्येक बाटलीत किती " + unit2 + " "
						+ name3 + " " + name5 + " " + name7 + "?<br>";

				// Generating correct and wrong answers
				int minus = a - c;
				int answer = minus / b;
				String correct_ans = "$" + answer + "$ " + unit + " <br>#$" + answer + "$ " + unit2 + " <br>";
				String wrong_ans1 = "$" + (answer + 1) + "$ " + unit + " <br>#$" + (answer + 1) + "$ " + unit2
						+ " <br>";
				String wrong_ans2 = "$" + (answer + 2) + "$ " + unit + " <br>#$" + (answer + 2) + "$ " + unit2
						+ " <br>";
				String wrong_ans3 = "$" + (answer - 1) + "$ " + unit + " <br>#$" + (answer - 1) + "$ " + unit2
						+ " <br>";

				// Generating solution
				String solution = "Ans : $" + answer + "$ " + unit + " <br>Let us assume, " + name2
						+ " filled in each container $=x$ " + unit + ".<br> " + name + " distributed $" + b
						+ "$ containers.<br> $\\therefore$ " + name2 + " in $" + b + "$ containers $=" + b
						+ "\\times x =" + b + "x$<br> Remaining " + name2 + " is $" + c + "$ " + unit
						+ "<br> Hence, total " + name2 + " given $=(" + b + "x+" + c + ")$ " + unit
						+ " $ . . . . (A)$<br> But total " + name2 + " prepared is $" + a + "$ " + unit
						+ " $. . . . (B)$<br>But as $A$ and $B$ are same,<br>  $\\therefore$ equating $A$ and $B$ we get<br> $"
						+ a + "=" + b + "x+" + c + "$<br> By subtracting $" + c + "$ from both sides we get<br> $" + a
						+ "-" + c + "= " + b + "x+" + c + "-" + c + "$<br> $\\therefore " + minus + "=" + b
						+ "x$<br> Now dividing both sides by $" + b + "$ we get <br> $x=" + answer + "$ " + unit
						+ "<br> $\\therefore " + answer + "$ " + unit
						+ " was filled in each container, is the answer.<br># उत्तर : $" + answer + "$ " + unit2
						+ " <br>प्रत्येक बाटलीत " + name8 + " " + name3 + " $=x$ " + unit2 + " मानू.<br>  " + name1
						+ " ने अशा $" + b + "$ बाटल्या मित्र-मैत्रिणींना दिल्या<br> $\\therefore " + b + "$ बाटल्यातील "
						+ name3 + " $=" + b + "\\times x =" + b + "x$<br> " + name9 + " " + name3 + " $" + c + "$ "
						+ unit2 + "<br> म्हणजे एकूण " + name3 + " $=(" + b + "x+" + c + ")$ " + unit2
						+ " $ . . . . (A)$<br> पण दिल्या नुसार तयार " + name10 + " एकूण " + name3 + " $" + a + "$ "
						+ unit2 + " $. . . . (B)$<br>आता $A$ आणि $B$ समान आहेत, <br> $\\therefore " + a + "=" + b + "x+"
						+ c + "$ असे मिळते. <br> दोन्ही बाजूतून $" + c + "$ वजा करून <br> $" + a + "-" + c + " =" + b
						+ "x+" + c + "-" + c + "$<br> $\\therefore " + minus + "=" + b + "x$<br> आता दोन्ही बाजुंना $"
						+ b + "$ ने भागून आपल्याला<br> $x=" + answer + "$ " + unit2
						+ " असे मिळते. <br> $\\therefore$ प्रत्येक बाटलीत $" + answer + "$ " + unit2 + " " + name3 + " "
						+ name5 + " हे उत्तर.<br>";

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
				row.createCell(12).setCellValue(60);
				row.createCell(13).setCellValue(2);
				row.createCell(15).setCellValue("siddhimanjarekar274@gmail.com");
				row.createCell(16).setCellValue(solution);
				row.createCell(18).setCellValue(124);

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
				continue; // Generate another question
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
}