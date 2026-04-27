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

public class Assign7_03040405_117 {

	public static Scanner sc = new Scanner(System.in);
	public static Random random = new Random();

	public static void main(String args[]) throws IOException, FileNotFoundException {

		String filename = "C:\\Siddhi\\Maths_Excelsheet\\VLab_03040405_117_Assign7_Siddhi.xlsx";
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
			int min = -9;
			int max = 20;
			int a = random.nextInt(max - min + 1) + min;
			a = checkIfZero(a);
			int b = random.nextInt(max - min + 1) + min;
			if (b == 0) {
				b = checkIfZero(b);
			} else {
				b = checkIfnp(a, b);
			}
			int c = random.nextInt(max - min + 1) + min;
			if (c == 0) {
				c = checkIfZero(c);
			} else {
				c = checkIfnp(a, c);
			}
			// variable
			String cht[] = { "x", "y", "z" };
			String cht1[] = { "u", "v", "w" };
			String cht2[] = { "a", "b", "c" };
			String cht3[] = { "p", "q", "r" };
			int chr = random.nextInt(cht.length);
			int chr1 = random.nextInt(cht.length);
			if (chr != chr1) {
				int chr2 = random.nextInt(cht.length);
				if (chr != chr2 && chr1 != chr2) {
					String ch1 = "";
					String ch2 = "";
					String ch3 = "";
					int ran = random.nextInt(4);
					switch (ran) {
					case 0:
						ch1 = cht[chr];
						ch2 = cht[chr1];
						ch3 = cht[chr2];
						break;
					case 1:
						ch1 = cht1[chr];
						ch2 = cht1[chr1];
						ch3 = cht1[chr2];
						break;
					case 2:
						ch1 = cht2[chr];
						ch2 = cht2[chr1];
						ch3 = cht2[chr2];
						break;
					case 3:
						ch1 = cht3[chr];
						ch2 = cht3[chr1];
						ch3 = cht3[chr2];
					}

					// powers
					int pow[] = { 1, 2, 3, 4 };
					int ex1 = random.nextInt(pow.length);
					int ex2 = random.nextInt(pow.length);
					if (ex1 != ex2) {
						int ex3 = random.nextInt(pow.length);
						if (ex1 != ex3 && ex2 != ex3) {
							int ex4 = random.nextInt(pow.length);
							if (ex1 != ex4 && ex2 != ex4 && ex3 != ex4) {
								int p1 = pow[ex1];
								int p2 = pow[ex2];
								int p3 = pow[ex3];
								int p4 = pow[ex4];
								int ab = a * b;
								int ac = a * c;
								int p1p1 = p1 + p1;
								int p2p3 = p2 + p3;
								int p2p4 = p2 + p4;
								String term1 = ab + power(ch1, p1p1) + power(ch2, p2p3) + power(ch3, p2) + ifn(ac) + ac
										+ power(ch1, p1) + power(ch2, p2p4) + power(ch3, p2);
								String term2 = if1(a) + power(ch1, p1) + power(ch2, p2);
								String term3 = if1(b) + power(ch1, p1) + power(ch2, p3) + power(ch3, p2);
								String term4 = if1(c) + power(ch2, p4) + power(ch3, p2);
								String Question = "One of the factors of $(" + term1 + ")$ is  $" + term2
										+ "$. What is the other factor?<br>#$(" + term1 + ")$ या राशीचा $" + term2
										+ "$ हा एक अवयव आहे. तर दिलेल्या पर्यायां पैकी दुसरा अवयव कोणता असेल?<br>";
								// Generate Correct answer
								String Correct_ans = "$(" + term3 + ifn(c) + term4 + ")$";
								// Generate wrong options

								String wrong_ans1 = "$(" + if1(b) + power(ch1, p1) + power(ch2, (p3 - 1))
										+ power(ch3, p2) + ifn(c) + if1(c) + power(ch2, p4) + power(ch3, p2) + ")$";

								String wrong_ans2 = "$(" + if1(b) + power(ch1, p1) + power(ch2, p3) + ifn(c) + if1(c)
										+ power(ch2, p4) + power(ch3, p2) + ")$";

								String wrong_ans3 = "$(" + if1(b) + power(ch2, p3) + power(ch3, p2) + ifn(c) + if1(c)
										+ power(ch2, p4) + ")$";

								String term5 = ab + power(ch1, p1p1) + power(ch2, (p2 + p3 - 1)) + power(ch3, p2)
										+ ifn(ac) + ac + power(ch1, p1) + power(ch2, p2p4) + power(ch3, p2);

								String term6 = ab + power(ch1, p1p1) + power(ch2, p2p3) + ifn(ac) + ac + power(ch1, p1)
										+ power(ch2, p2p4) + power(ch3, p2);

								String term7 = ab + power(ch1, p1) + power(ch2, p2p3) + power(ch3, p2) + ifn(ac) + ac
										+ power(ch1, p1) + power(ch2, p2p4);

								String Solution = "​Ans :  " + Correct_ans + "  <br>To find out other factor of $("
										+ term1 + ")$,<br> Let us take the multiplication of given factor $" + term2
										+ "$ with all options, one by one.<br>By taking these multiplications<br>For option "
										+ wrong_ans1 + " we get<br>$\\therefore " + term2 + "  \\times (" + if1(b)
										+ power(ch1, p1) + power(ch2, (p3 - 1)) + power(ch3, p2) + ifn(c) + if1(c)
										+ power(ch2, p4) + power(ch3, p2) + ") = " + term5
										+ " \\ne$  the given expression.<br>For option " + wrong_ans2
										+ " we get,<br>$\\therefore " + term2 + "  \\times (" + if1(b) + power(ch1, p1)
										+ power(ch2, p3) + ifn(c) + if1(c) + power(ch2, p4) + power(ch3, p2) + ") = "
										+ term6 + " \\ne$  the given expression.<br>For option" + wrong_ans3
										+ " we get,<br>$\\therefore " + term2 + "  \\times (" + if1(b) + power(ch2, p3)
										+ power(ch3, p2) + ifn(c) + if1(c) + power(ch2, p4) + ") = " + term7
										+ " \\ne$  the given expression.<br>For option " + Correct_ans
										+ " we get,<br>$\\therefore " + term2 + "  \\times (" + term3 + ifn(c) + term4
										+ ") = " + term1
										+ " =$ given expression.<br>$\\therefore$ required other  factor is "
										+ Correct_ans + " is the answer<br>#उत्तर :  " + Correct_ans + "  <br>$("
										+ term1
										+ ")$ या दिलेल्या राशीचा दुसरा अवयव शोधण्या साठी<br> आपण दिलेल्या प्रत्येक पर्यायाचा $"
										+ term2 + "$ या दिलेल्या अवयव बरोबर गुणाकार करू. <br>पर्याय " + wrong_ans1
										+ " साठी आपल्याला <br>$ " + term2 + "  \\times (" + if1(b) + power(ch1, p1)
										+ power(ch2, (p3 - 1)) + power(ch3, p2) + ifn(c) + if1(c) + power(ch2, p4)
										+ power(ch3, p2) + ") = " + term5
										+ "$ असे मिळते आणि ही राशी दिलेल्या राशी पेक्षा वेगळी आहे.<br>पर्याय  "
										+ wrong_ans2 + " साठी आपल्याला <br>$ " + term2 + "  \\times (" + if1(b)
										+ power(ch1, p1) + power(ch2, p3) + ifn(c) + if1(c) + power(ch2, p4)
										+ power(ch3, p2) + ") = " + term6
										+ " $  असे मिळते आणि ही राशी दिलेल्या राशी पेक्षा वेगळी आहे.<br>पर्याय  "
										+ wrong_ans3 + " साठी आपल्याला <br>$ " + term2 + "  \\times (" + if1(b)
										+ power(ch2, p3) + power(ch3, p2) + ifn(c) + if1(c) + power(ch2, p4) + ") = "
										+ term7 + " $ असे मिळते आणि ही राशी दिलेल्या राशी पेक्षा वेगळी आहे.<br>पर्याय  "
										+ Correct_ans + " साठी आपल्याला <br>$ " + term2 + "  \\times (" + term3 + ifn(c)
										+ term4 + ") = " + term1
										+ " $ असे मिळते आणि ही राशी दिलेल्या राशी इतकीच आहे.<br>"
										+ "$\\therefore$ दुसरा अवयव " + Correct_ans + " आहे, हे उत्तर.<br> ";

								XSSFRow row = sheet1.createRow(i);
								row.createCell(0).setCellValue(i);
								row.createCell(1).setCellValue("Text");
								row.createCell(2).setCellValue(1);
								row.createCell(3).setCellValue("03040405");
								row.createCell(4).setCellValue(Question);
								row.createCell(5).setCellValue(Correct_ans + "<br>");
								row.createCell(9).setCellValue(wrong_ans1 + "<br>");
								row.createCell(10).setCellValue(wrong_ans2 + "<br>");
								row.createCell(11).setCellValue(wrong_ans3 + "<br>");
								row.createCell(12).setCellValue(120);
								row.createCell(13).setCellValue(3);
//                  			row.createCell(14).setCellValue(" ");
								row.createCell(15).setCellValue("siddhimanjarekar274@gmail.com");
								row.createCell(16).setCellValue(Solution);
//                  			row.createCell(17).setCellValue(" ");
								row.createCell(18).setCellValue(117);
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
										|| wrong_ans1 == wrong_ans2 || wrong_ans1 == wrong_ans3
										|| wrong_ans2 == wrong_ans3) {
									System.out.println("duplicate" + i);
									i--;
								}
							} else {
								continue;
							}
						} else {
							continue;
						}
					} else {
						continue;
					}
				} else {
					continue;
				}
			} else {
				continue;
			}

			i++;
		} while (i < q + 1);
		int rowTotal = sheet1.getLastRowNum();
//      System.out.println(rowTotal);
		XSSFRow row1 = sheet1.createRow((short) rowTotal + 1);
		row1.createCell(0).setCellValue("");

		// Writing data to the file
		FileOutputStream fileout = new FileOutputStream(filename);
		workbook.write(fileout);
		fileout.close();

		System.out.println("file created");
	}

	public static String power(String ch, int pow) {
		if (pow != 0 && pow != 1) {
			return ch + "^" + pow;
		} else if (pow == 1) {
			return ch;
		} else if (pow == 0) {
			return " ";
		}
		return null;
	}

	public static String if1(int a) {
		if (a == 1)
			return "";
		else if (a == -1)
			return "-";
		else
			return String.valueOf(a);
	}

	public static String ifn(int a) {
		if (a < 0) {
			return "";
		} else {
			return "+";
		}
	}

	public static int checkIfZero(int a) {
		if (a != 0) {
			return a;
		} else {
			do {
				a = random.nextInt(10);
			} while (a == 0);
		}
		return a;
	}

	public static int checkIfnp(int a, int b) {
		if (a != (-b)) {
			return b;
		} else {
			do {
				b = random.nextInt(10);
			} while (a == (-b));
		}
		return b;
	}
}
