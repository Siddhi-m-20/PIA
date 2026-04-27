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

public class Assign9_03040405_120 {
	public static Scanner sc = new Scanner(System.in);
	public static Random random = new Random();

	public static void main(String args[]) throws IOException, FileNotFoundException {
		String filename = "C:\\Siddhi\\Maths_Excelsheet\\VML_03040405_120_Assign9_Siddhi.xlsx";
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

		int mapsize, mapsizeafter;
		HashMap<String, Integer> map = new HashMap<String, Integer>();
		int q = sc.nextInt();
		int i = 1;
		do {
			Random r = new Random();
			int min = 1;
			int max = 20;
			int a = random.nextInt(max - min + 1) + min;
			int b = random.nextInt(max - min + 1) + min;
			int asqr = a * a;
			int bsqr = b * b;
			int[] p1 = { 2, 4 };
			int p = p1[random.nextInt(p1.length)];
			// variable
			String cht[] = { "u", "v", "w", "x", "y", "z" };
			int chr = random.nextInt(cht.length);
			int chr1 = random.nextInt(cht.length);
			if (chr != chr1 && p == 2) {
				String ch = cht[chr];
				String ch1 = cht[chr1];
				String Question = "Find factors of following expression,<br> $" + hide1a(asqr, ch) + "^2"
						+ "- \\left(\\dfrac{" + bsqr + "}{" + ch1 + "^" + p
						+ "}\\right)$.<br># खालील राशीचे अवयव शोधा,<br>$" + hide1a(asqr, ch) + "^2"
						+ "- \\left(\\dfrac{" + bsqr + "}{" + ch1 + "^" + p + "}\\right)$.<br>";
				String Correct_ans = " $\\left(" + hide1a(a, ch) + "+ \\dfrac{" + b + "}{" + ch1 + "}\\right)\\left("
						+ hide1a(a, ch) + "- \\dfrac{" + b + "}{" + ch1 + "}\\right)$ <br>";
				String wrong_ans1 = " $\\left(" + hide1a(a, ch) + "+ \\dfrac{" + b + "}{" + ch1 + "}\\right)\\left("
						+ hide1a(a, ch) + "+ \\dfrac{" + b + "}{" + ch1 + "}\\right)$ <br>";
				String wrong_ans2 = " $\\left(" + hide1a(a, ch) + "- \\dfrac{" + b + "}{" + ch1 + "}\\right)\\left("
						+ hide1a(a, ch) + "- \\dfrac{" + b + "}{" + ch1 + "}\\right)$ <br>";
				String wrong_ans3 = " $\\left(" + wrong_option(a) + ch + "- \\dfrac{" + b + "}{" + ch1 + "}\\right)\\left("
						+ wrong_option(a) + ch + "- \\dfrac{" + b + "}{" + ch1 + "}\\right)$ <br>";
				String wrong_ans4 = " $\\left(" + hide1a(a, ch) + "+ \\dfrac{" + wrong_option(b) + "}{" + ch1
						+ "}\\right)\\left(" + hide1a(a, ch) + "+ \\dfrac{" + wrong_option(b) + "}{" + ch1
						+ "}\\right)$ <br>";
				String wrong_ans5 = " $\\left(" + (-a) + ch + "+ \\dfrac{" + b + "}{" + ch1 + "}\\right)\\left("
						+ hide1a(a, ch) + "- \\dfrac{" + b + "}{" + ch1 + "}\\right)$ <br>";
				String Solution = "Ans :   $\\left(" + hide1a(a, ch) + "+ \\dfrac{" + b + "}{" + ch1 + "}\\right)\\left("
						+ hide1a(a, ch) + "- \\dfrac{" + b + "}{" + ch1 + "}\\right)$ <br>"
						+ "We can see that the given binomial is the subtraction of two perfect square terms, and<br>"
						+ "we also know that ​​​​$a^2 - b^2 = (a + b)(a - b)$.<br>"
						+ "This shows that ​​​$(a + b)$ and $(a - b)$ are the factors of $a^2 - b^2$.<br>"
						+ "From this we can write,<br>$" + hide1a(asqr, ch) + "^2" + "- \\left(\\dfrac{" + bsqr + "}{"
						+ ch1 + "^" + p + "}\\right)$<br>" + "$=\\left(" + hide1a(a, ch) + "\\right)^2 - \\left( \\dfrac{" + b + "}{"
						+ ch1 + "} \\right)^2$<br>" + "$=\\left(" + hide1a(a, ch) + "+ \\dfrac{" + b + "}{" + ch1
						+ "}\\right)\\left(" + hide1a(a, ch) + "- \\dfrac{" + b + "}{" + ch1 + "}\\right)$ <br>"
						+ "$\\therefore \\left(" + hide1a(a, ch) + "+ \\dfrac{" + b + "}{" + ch1 + "}\\right)\\left("
						+ hide1a(a, ch) + "- \\dfrac{" + b + "}{" + ch1
						+ "}\\right)$ are the required factors for the given binomial.<br>" + "#उत्तर :  $\\left("
						+ hide1a(a, ch) + "+ \\dfrac{" + b + "}{" + ch1 + "}\\right)\\left(" + hide1a(a, ch)
						+ "- \\dfrac{" + b + "}{" + ch1 + "}\\right)$  <br>"
						+ "आपण हे पाहू शकतो की दिलेली द्विपदी ही दोन पूर्ण वर्ग पदांची वजाबाकी आहे आणि<br>"
						+ "आपल्याला हे पण माहिती आहे की ​​​​$a^2 - b^2 = (a + b)(a - b)$.<br>"
						+ "यावरून आपण हे सांगू शकतो की ​$(a + b)$ आणि $(a - b)$ हे $a^2 - b^2$ याचे अवयव आहेत.<br>"
						+ "म्हणून आपण खालील प्रमाणे लिहू शकतो,<br>$" + hide1a(asqr, ch) + "^2" + "- \\left(\\dfrac{"
						+ bsqr + "}{" + ch1 + "^" + p + "}\\right)$<br>" + "$=\\left(" + hide1a(a, ch)
						+ "\\right)^2 - \\left( \\dfrac{" + b + "}{" + ch1 + "} \\right)^2$<br>" + "$=\\left(" + hide1a(a, ch)
						+ "+ \\dfrac{" + b + "}{" + ch1 + "}\\right)\\left(" + hide1a(a, ch) + "- \\dfrac{" + b
						+ "}{" + ch1 + "}\\right)$ <br>" + "$\\therefore \\left(" + hide1a(a, ch) + "+ \\dfrac{" + b
						+ "}{" + ch1 + "}\\right)\\left(" + hide1a(a, ch) + "- \\dfrac{" + b + "}{" + ch1
						+ "}\\right)$ हे दिलेल्या द्विपदीचे अवयव आहेत हे उत्तर.<br>";

				Vector<String> vec = new Vector<String>();
				vec.add(wrong_ans1);
				vec.add(wrong_ans2);
				vec.add(wrong_ans3);
				vec.add(wrong_ans4);
				vec.add(wrong_ans5);
				int wrg = r.nextInt(5);
				HashSet<String> set = new HashSet<String>();
				for (int j = 0; j < vec.size(); j++) {
					set.add(vec.elementAt(j));
				}
				XSSFRow row = sheet1.createRow(i);
				row.createCell(0).setCellValue(i);
				row.createCell(1).setCellValue("Text");
				row.createCell(2).setCellValue(1);
				row.createCell(3).setCellValue("03040405");
				row.createCell(4).setCellValue(Question);
				row.createCell(5).setCellValue(Correct_ans);
				row.createCell(9).setCellValue(vec.elementAt(wrg));
				vec.removeElementAt(wrg);
				wrg = r.nextInt(4);
				row.createCell(10).setCellValue(vec.elementAt(wrg));
				vec.removeElementAt(wrg);
				wrg = r.nextInt(3);
				row.createCell(11).setCellValue(vec.elementAt(wrg));
				vec.removeElementAt(wrg);
				row.createCell(12).setCellValue(90);
				row.createCell(13).setCellValue(2);
//			row.createCell(14).setCellValue(" ");
				row.createCell(15).setCellValue("siddhimanjarekar274@gmail.com");
				row.createCell(16).setCellValue(Solution);
//			row.createCell(17).setCellValue(" ");
				row.createCell(18).setCellValue(120);

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
						|| Correct_ans == wrong_ans4 || Correct_ans == wrong_ans5 || wrong_ans1 == wrong_ans2
						|| wrong_ans1 == wrong_ans3 || wrong_ans1 == wrong_ans4 || wrong_ans1 == wrong_ans5
						|| wrong_ans2 == wrong_ans3 || wrong_ans2 == wrong_ans4 || wrong_ans2 == wrong_ans5
						|| wrong_ans3 == wrong_ans4 || wrong_ans3 == wrong_ans4 || wrong_ans4 == wrong_ans5) {
					System.out.println("duplicate" + i);
					i--;
				}
			} else if (chr != chr1 && p == 4) {
				String ch = cht[chr];
				String ch1 = cht[chr1];
				String Question = "Find factors of following expression,<br> $" + hide1a(asqr, ch) + "^2"
						+ "- \\left(\\dfrac{" + bsqr + "}{" + ch1 + "^" + p
						+ "}\\right)$.<br># खालील राशीचे अवयव शोधा,<br>$" + hide1a(asqr, ch) + "^2"
						+ "- \\left(\\dfrac{" + bsqr + "}{" + ch1 + "^" + p + "}\\right)$.<br>";
				String Correct_ans = " $\\left(" + hide1a(a, ch) + "+ \\dfrac{" + b + "}{" + ch1 + "^2" + "}\\right)\\left("
						+ hide1a(a, ch) + "- \\dfrac{" + b + "}{" + ch1 + "^2" + "}\\right)$ <br>";
				String wrong_ans1 = " $\\left(" + hide1a(a, ch) + "+ \\dfrac{" + b + "}{" + ch1 + "^2" + "}\\right)\\left("
						+ hide1a(a, ch) + "+ \\dfrac{" + b + "}{" + ch1 + "^2" + "}\\right)$ <br>";
				String wrong_ans2 = " $\\left(" + hide1a(a, ch) + "- \\dfrac{" + b + "}{" + ch1 + "^2" + "}\\right)\\left("
						+ hide1a(a, ch) + "- \\dfrac{" + b + "}{" + ch1 + "^2" + "}\\right)$ <br>";
				String wrong_ans3 = " $\\left(" + wrong_option(a) + ch + "- \\dfrac{" + b + "}{" + ch1 + "^2"
						+ "}\\right)\\left(" + wrong_option(a) + ch + "- \\dfrac{" + b + "}{" + ch1 + "^2"
						+ "}\\right)$ <br>";
				String wrong_ans4 = " $\\left(" + hide1a(a, ch) + "+ \\dfrac{" + wrong_option(b) + "}{" + ch1 + "^2"
						+ "}\\right)\\left(" + hide1a(a, ch) + "+ \\dfrac{" + wrong_option(b) + "}{" + ch1 + "^2"
						+ "}\\right)$ <br>";
				String wrong_ans5 = " $\\left(" + (-a) + ch + "+ \\dfrac{" + b + "}{" + ch1 + "^2" + "}\\right)\\left("
						+ hide1a(a, ch) + "- \\dfrac{" + b + "}{" + ch1 + "^2" + "}\\right)$ <br>";
				String Solution = "Ans :   $\\left(" + hide1a(a, ch) + "+ \\dfrac{" + b + "}{" + ch1 + "^2"
						+ "}\\right)\\left(" + hide1a(a, ch) + "- \\dfrac{" + b + "}{" + ch1 + "^2"
						+ "}\\right)$ <br>"
						+ "We can see that the given binomial is the subtraction of two perfect square terms, and<br>"
						+ "we also know that ​​​​$a^2 - b^2 = (a + b)(a - b)$.<br>"
						+ "This shows that​ ​​​$(a + b)$ and $(a - b)$ are the factors of $a^2 - b^2$.<br>"
						+ "​From this we can write,<br>$" + hide1a(asqr, ch) + "^2" + "- \\left(\\dfrac{" + bsqr + "}{"
						+ ch1 + "^" + p + "}\\right)$<br>" + "$=\\left(" + hide1a(a, ch) + "\\right)^2 - \\left( \\dfrac{" + b + "}{"
						+ ch1 + "^2} \\right)^2$<br>" + "$=\\left(" + hide1a(a, ch) + "+ \\dfrac{" + b + "}{" + ch1
						+ "^2}\\right)\\left(" + hide1a(a, ch) + "- \\dfrac{" + b + "}{" + ch1 + "^2}\\right)$ <br>"
						+ "$\\therefore \\left(" + hide1a(a, ch) + "+ \\dfrac{" + b + "}{" + ch1 + "^2}\\right)\\left("
						+ hide1a(a, ch) + "- \\dfrac{" + b + "}{" + ch1
						+ "^2}\\right)$  are the required factors for the given binomial.<br>" + "#उत्तर :  $\\left("
						+ hide1a(a, ch) + "+ \\dfrac{" + b + "}{" + ch1 + "^2" + "}\\right)\\left(" + hide1a(a, ch)
						+ "- \\dfrac{" + b + "}{" + ch1 + "^2" + "}\\right)$  <br>"
						+ "आपण हे पाहू शकतो की दिलेली द्विपदी ही दोन पूर्ण वर्ग पदांची वजाबाकी आहे आणि<br>"
						+ "आपल्याला हे पण माहिती आहे की ​​​​$a^2 - b^2 = (a + b)(a - b)$.<br>"
						+ "यावरून आपण हे सांगू शकतो की ​$(a + b)$ आणि $(a - b)$ हे $a^2 - b^2$ याचे अवयव आहेत.<br>"
						+ "म्हणून आपण खालील प्रमाणे लिहू शकतो,<br>$" + hide1a(asqr, ch) + "^2" + "- \\left(\\dfrac{"
						+ bsqr + "}{" + ch1 + "^" + p + "}\\right)$<br>" + "$=\\left(" + hide1a(a, ch)
						+ "\\right)^2 - \\left( \\dfrac{" + b + "}{" + ch1 + "^2} \\right)^2$<br>" + "$=\\left(" + hide1a(a, ch)
						+ "+ \\dfrac{" + b + "}{" + ch1 + "^2}\\right)\\left(" + hide1a(a, ch) + "- \\dfrac{"
						+ b + "}{" + ch1 + "^2}\\right)$ <br>" + "$\\therefore \\left(" + hide1a(a, ch) + "+ \\dfrac{"
						+ b + "}{" + ch1 + "^2}\\right)\\left(" + hide1a(a, ch) + "- \\dfrac{" + b + "}{" + ch1
						+ "^2}\\right)$ हे दिलेल्या द्विपदीचे अवयव आहेत हे उत्तर.<br>";
				Vector<String> vec = new Vector<String>();
				vec.add(wrong_ans1);
				vec.add(wrong_ans2);
				vec.add(wrong_ans3);
				vec.add(wrong_ans4);
				vec.add(wrong_ans5);
				int wrg = r.nextInt(5);
				HashSet<String> set = new HashSet<String>();
				for (int j = 0; j < vec.size(); j++) {
					set.add(vec.elementAt(j));
				}
				XSSFRow row = sheet1.createRow(i);
				row.createCell(0).setCellValue(i);
				row.createCell(1).setCellValue("Text");
				row.createCell(2).setCellValue(1);
				row.createCell(3).setCellValue("03040405");
				row.createCell(4).setCellValue(Question);
				row.createCell(5).setCellValue(Correct_ans);
				row.createCell(9).setCellValue(vec.elementAt(wrg));
				vec.removeElementAt(wrg);
				wrg = r.nextInt(4);
				row.createCell(10).setCellValue(vec.elementAt(wrg));
				vec.removeElementAt(wrg);
				wrg = r.nextInt(3);
				row.createCell(11).setCellValue(vec.elementAt(wrg));
				vec.removeElementAt(wrg);
				row.createCell(12).setCellValue(90);
				row.createCell(13).setCellValue(2);
//			row.createCell(14).setCellValue(" ");
				row.createCell(15).setCellValue("siddhimanjarekar274@gmail.com");
				row.createCell(16).setCellValue(Solution);
//			row.createCell(17).setCellValue(" ");
				row.createCell(18).setCellValue(120);

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
						|| Correct_ans == wrong_ans4 || Correct_ans == wrong_ans5 || wrong_ans1 == wrong_ans2
						|| wrong_ans1 == wrong_ans3 || wrong_ans1 == wrong_ans4 || wrong_ans1 == wrong_ans5
						|| wrong_ans2 == wrong_ans3 || wrong_ans2 == wrong_ans4 || wrong_ans2 == wrong_ans5
						|| wrong_ans3 == wrong_ans4 || wrong_ans3 == wrong_ans4 || wrong_ans4 == wrong_ans5) {
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

	public static String hide1a(int a, String chr) {
		if (a != 1 && a != 0 && a != -1)
			return a + chr;
		else if (a == 1)
			return chr;
		else if (a == -1) {
			return String.valueOf("-" + chr);
		}
		return null;
	}

	public static String wrong_option(int a) {
		if (a == 1) {
			return String.valueOf(2);
		} else if (a == 2) {
			return String.valueOf(8);
		} else if (a == 3) {
			return String.valueOf(6);
		} else if (a == 4) {
			return String.valueOf(8);
		} else if (a == 5) {
			return String.valueOf(10);
		} else if (a == 6) {
			return String.valueOf(4);
		} else if (a == 7) {
			return String.valueOf(3);
		} else if (a == 8) {
			return String.valueOf(2);
		} else if (a == 9) {
			return String.valueOf(11);
		} else if (a == 10) {
			return String.valueOf(5);
		} else if (a == 11) {
			return String.valueOf(19);
		} else if (a == 12) {
			return String.valueOf(18);
		} else if (a == 13) {
			return String.valueOf(17);
		} else if (a == 14) {
			return String.valueOf(16);
		} else if (a == 15) {
			return String.valueOf(20);
		} else if (a == 16) {
			return String.valueOf(14);
		} else if (a == 17) {
			return String.valueOf(13);
		} else if (a == 18) {
			return String.valueOf(12);
		} else if (a == 19) {
			return String.valueOf(21);
		} else if (a == 20) {
			return String.valueOf(15);
		} else {
			return String.valueOf(a);
		}
	}
}
