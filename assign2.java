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
import java.lang.Math;

public class assign2 {

	public static Scanner sc = new Scanner(System.in);
	public static Random random = new Random();

	public static void main(String[] args) throws IOException, FileNotFoundException {
		String filename = "C:\\Siddhi\\Maths_Excelsheet\\VLab_304050404_109_1_Assign2_Siddhi.xlsx";
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
		int i = 1;
		do {

			// coeffients ranges
			int min1 = -9;
			int max1 = 9;
			int a = random.nextInt(max1 - min1 + 1) + min1;
			a = checkIfZero(a);
			int b = random.nextInt(max1 - min1 + 1) + min1;
			b = checkIfZero(b);
			int c = random.nextInt(max1 - min1 + 1) + min1;
			c = checkIfZero(c);
			int d = random.nextInt(max1 - min1 + 1) + min1;
			d = checkIfZero(d);
			int e = random.nextInt(max1 - min1 + 1) + min1;
			e = checkIfZero(e);

			// power
			int p1 = 2;
			int p2 = 1;
			int p3 = 2;
			int p4 = 1;
			// variable
			String cht[] = { "x", "y", "z" };
			int chr = random.nextInt(cht.length);
			String ch = cht[chr];

			String Question = "A square shaped land with each side $" + hidecoepow(a, ch, p1) + plusign(b)
					+ hidecoepow(b, ch, p2) + plusign(c) + c + "$ is to be divided into equal plots of the area $"
					+ hidecoepow(d, ch, p3) + plusign(e) + hidecoepow(e, ch, p4)
					+ "$. How many such plots are possible and is there any remaining area of the land.<br >#​​" + "$"
					+ hidecoepow(a, ch, p1) + plusign(b) + hidecoepow(b, ch, p2) + plusign(c) + c
					+ "$ या मापाची बाजू असलेली एका चौरसाकार जमिनीचे  $" + hidecoepow(d, ch, p3) + plusign(e)
					+ hidecoepow(e, ch, p4)
					+ "$ या समान मापाचे तुकडे करायचे आहेत, तर असे किती तुकडे तयार होतील आणि किती जमीन शिल्लक राहील?<br>";

			// multiplication of 1st term to whole equation
			int m1aa = a * a;// 1st coe
			int m1p1 = p1 + p1;// 1st pow
			int m1ab = a * b;// 2nd coe
			int m1p2 = p1 + p2;// 2nd pow
			int m1ac = a * c;// 3rd coe
			int m1p3 = p1;
			String m1 = hidecoepow(m1aa, ch, m1p1) + plusign(m1ab) + hidecoepow(m1ab, ch, m1p2) + plusign(m1ac)
					+ hidecoepow(m1ac, ch, m1p3);

			// multiplication of 2nd term to whole equation
			int m2ba = b * a;// 1st coe
			int m2p1 = p2 + p1;// 1st pow
			int m2bb = b * b;// 2nd coe
			int m2p2 = p2 + p2;// 2nd pow
			int m2bc = b * c;// 3rd coe
			int m2p3 = p2;
			String m2 = plusign(m2ba) + hidecoepow(m2ba, ch, m2p1) + plusign(m2bb) + hidecoepow(m2bb, ch, m2p2)
					+ plusign(m2bc) + hidecoepow(m2bc, ch, m2p3);
			// multiplication of 3rd term to whole equation
			int m3ca = c * a;// 1st coe
			int m3p1 = p1;// 1st pow
			int m3cb = c * b;// 2nd coe
			int m3p2 = p2;// 2nd pow
			int m3cc = c * c;// 3rd coe
			String m3 = plusign(m3ca) + hidecoepow(m3ca, ch, m3p1) + plusign(m3cb) + hidecoepow(m3cb, ch, m3p2)
					+ plusign(m3cc) + m3cc;

			int pow3 = m1ab + m2ba;
			int pow2 = m1ac + m2bb + m3ca;
			int pow1 = m2bc + m3cb;
			String st3 = hidecoepow(m1aa, ch, m1p1) + common_sign2(m1ab, m2ba) + "(" + plusign2(m1ab, m2ba) + ")" + ch
					+ "^" + m1p2 + common_sign3(m1ac, m2bb, m3ca) + "(" + plusign3(m1ac, m2bb, m3ca) + ")" + ch + "^"
					+ m1p3 + common_sign2(m2bc, m3cb) + "(" + plusign2(m2bc, m3cb) + ")" + ch + plusign(m3cc) + m3cc;

			// Resultant Square of side
			String Eqsq = hidecoepow(m1aa, ch, m1p1) + plusign(pow3) + hidecoepow(pow3, ch, m1p2) + plusign(pow2)
					+ hidecoepow(pow2, ch, m1p3) + plusign(pow1) + hidecoepow(pow1, ch, m2p3) + plusign(m3cc) + m3cc;

			if (m1aa % d == 0 && m1aa > 0 && d > 0 && a > 0) {
				int s1 = m1aa / d;
				int ex1 = m1p1 - p3;
				int s1d = s1 * d;
				int ex1p3 = ex1 + p3;
				int s1e = s1 * e;
				int ex1p4 = ex1 + p4;
				String sub1 = hidecoepow(s1d, ch, ex1p3) + plusign(s1e) + hidecoepow(s1e, ch, ex1p4);// 1st -string

				int ans1v = pow3 - s1e;
				String ans1 = hidecoepow(ans1v, ch, ex1p4) + plusign(pow2) + hidecoepow(pow2, ch, m1p3) + plusign(pow1)
						+ hidecoepow(pow1, ch, m2p3) + plusign(m3cc) + m3cc;// 1st answer string
				if (ans1v % d == 0) {
					int s2 = ans1v / d;
					int ex2 = ex1p4 - p3;
					int s2d = s2 * d;
					int ex2p3 = ex2 + p3;
					int s2e = s2 * e;
					int ex2p4 = ex2 + p4;
					String sub2 = hidecoepow(s2d, ch, ex2p3) + plusign(s2e) + hidecoepow(s2e, ch, ex2p4);// 2nd -string
					int ans2v = pow2 - s2e;
					String ans2 = hidecoepow(ans2v, ch, m1p3) + plusign(pow1) + hidecoepow(pow1, ch, m2p3)
							+ plusign(m3cc) + m3cc;// 2nd answer string
					if (ans2v % d == 0) {
						int s3 = ans2v / d;
						if (s3 != 0) {
							int ex3 = m1p3 - p3;
							int s3d = s3 * d;
							int ex3p3 = ex3 + p3;
							int s3e = s3 * e;
							int ex3p4 = ex3 + p4;
							String sub3 = hidecoepow(s3d, ch, ex3p3) + plusign(s3e) + hidecoepow(s3e, ch, ex3p4);// 3rd
																													// -string
							int ans3v = pow1 - s3e;
							String remaining_part = "$" + hidecoepow(ans3v, ch, m2p3) + plusign(m3cc) + m3cc + "$";// 2nd
							if (ans3v < 500 && ans3v > -200) {
								String no_of_parts = hidecoepow(s1, ch, ex1) + plusign(s2) + hidecoepow(s2, ch, ex2)
										+ plusign(s3) + hidecoepow(s3, ch, ex3);
								if (s1 > 0 && ans3v > 0) {
									String sol = "No. of plots  " + "$" + no_of_parts + "$  and remaining area   "
											+ remaining_part + "   Sq. units<br>#जमिनीचे सारख्या मापाचे तुकडे   $"
											+ no_of_parts + "$  आणि शिल्लक जमीन   " + remaining_part
											+ "   चौ. एकक <br>";
									String wrong_ans1 = "No. of plots $" + hidecoepow(s2, ch, ex2) + plusign(s3)
											+ hidecoepow(s3, ch, ex3) + plusign(s1) + hidecoepow(s1, ch, ex1)
											+ "$  and remaining area   $" + hidecoepow(ans2v, ch, m2p3) + plusign(m3cc)
											+ m3cc + "$  Sq. units<br>#जमिनीचे सारख्या मापाचे तुकडे  $"
											+ hidecoepow(s2, ch, ex2) + plusign(s2) + hidecoepow(s3, ch, ex3)
											+ plusign(s3) + hidecoepow(s1, ch, ex1) + "$  आणि शिल्लक जमीन   $"
											+ hidecoepow(ans2v, ch, m2p3) + plusign(m3cc) + m3cc + "$   चौ. एकक <br>";

									String wrong_ans2 = "No. of plots  $" + hidecoepow(s3, ch, ex3) + plusign(s2)
											+ hidecoepow(s2, ch, ex2) + plusign(s3) + hidecoepow(s1, ch, ex1)
											+ "$  and remaining area    $" + hidecoepow(ans1v, ch, p1) + plusign(m3cc)
											+ m3cc + "$   Sq. units<br>#जमिनीचे सारख्या मापाचे तुकडे    $"
											+ hidecoepow(s3, ch, ex3) + plusign(s2) + hidecoepow(s2, ch, ex2)
											+ plusign(s3) + hidecoepow(s1, ch, ex1) + "$  आणि शिल्लक जमीन   $"
											+ hidecoepow(ans1v, ch, p1) + plusign(m3cc) + m3cc + "$    चौ. एकक <br>";
									String wrong_ans3 = "No. of plots  $" + hidecoepow(s3, ch, ex3) + plusign(s1)
											+ hidecoepow(s1, ch, ex1) + plusign(s2) + hidecoepow(s2, ch, ex2)
											+ "$  and remaining area   $" + hidecoepow(ans2v, ch, p2) + plusign(m3cc)
											+ m3cc + "$   Sq. units<br>#जमिनीचे सारख्या मापाचे तुकडे   $"
											+ hidecoepow(s3, ch, ex3) + plusign(s2) + hidecoepow(s1, ch, ex1)
											+ plusign(s3) + hidecoepow(s2, ch, ex2) + "$  आणि शिल्लक जमीन   $"
											+ hidecoepow(ans2v, ch, p2) + plusign(m3cc) + m3cc + "$    चौ. एकक <br>";

									String soleng = "Ans : No. of plots $" + no_of_parts + "$   and remaining area "
											+ remaining_part
											+ " Sq. units<br>As area of the land to be divided into plots is of square shape, it's side are always equal.<br>"
											+ "Therefore, for finding the area of the shape, we need to multiply the given expression by itself.<br>"
											+ "Therefore the area to be divided in equal parts <br>$= ("
											+ hidecoepow(a, ch, p1) + plusign(b) + hidecoepow(b, ch, p2) + plusign(c)
											+ c + ")\\times(" + hidecoepow(a, ch, p1) + plusign(b)
											+ hidecoepow(b, ch, p2) + plusign(c) + c + ")$<br>$= " + m1 + m2 + m3
											+ "$<br>$= " + st3 + "$<br>$= " + Eqsq
											+ " . . . . .(i)$ <br>To divide this area into equal number of plots of area $"
											+ hidecoepow(d, ch, p3) + plusign(e) + hidecoepow(e, ch, p4)
											+ "$, we need to divide equation $(i)$ by $" + hidecoepow(d, ch, p3)
											+ plusign(e) + hidecoepow(e, ch, p4)
											+ "$<br>Set up the division according to the standard form of long division with dividend and divisor written as shown.<br> "
											+ "$ \\begin{array}{r}" + no_of_parts
											+ "\\phantom{0000000000000} \\\\\\phantom{00000000} "
											+ hidecoepow(d, ch, p3) + plusign(e) + hidecoepow(e, ch, p4)
											+ "\\big)\\overline{" + Eqsq + "} \\\\" + sub1
											+ "\\phantom{000000000000000} \\\\\\underline{-\\phantom{000} -\\phantom{0000000000000000000} }\\\\"
											+ ans1 + "\\phantom{0} \\\\" + sub2
											+ "\\phantom{0000000000}\\\\\\underline{-\\phantom{000} -\\phantom{00000000000000000}} \\\\ "
											+ ans2 + "\\phantom{} \\\\" + sub3
											+ "\\phantom{000}\\\\\\underline{-\\phantom{000} -\\phantom{0000000} } \\\\"
											+ hidecoepow(ans3v, ch, m2p3) + plusign(m3cc) + m3cc
											+ "\\phantom{} \\end{array}$<br><br>" + "By dividing we get quotient $="
											+ hidecoepow(s1, ch, ex1) + plusign(s2) + hidecoepow(s2, ch, ex2)
											+ plusign(s3) + hidecoepow(s3, ch, ex3) + "$ and remainder $= "
											+ hidecoepow(ans3v, ch, m2p3) + plusign(m3cc) + m3cc
											+ "$<br>$\\Rightarrow$ No. of plots $" + hidecoepow(s1, ch, ex1)
											+ plusign(s2) + hidecoepow(s2, ch, ex2) + plusign(s3)
											+ hidecoepow(s3, ch, ex3) + "$  and remaining area $"
											+ hidecoepow(ans3v, ch, m2p3) + plusign(m3cc) + m3cc
											+ "$ Sq. units, is the answer<br>";
									String solmar = "#उत्तर : जमिनीचे सारख्या मापाचे तुकडे   $" + no_of_parts
											+ "$  आणि शिल्लक जमीन " + remaining_part
											+ "  चौ. एकक <br>ज्या जमिनीचे सारख्या मापाचे तुकडे करायचे आहेत तिचा आकार चौरस आहे, म्हणजेच तिची लांबी आणि रुंदी समान आहे.<br>"
											+ "अशा आकाराचे क्षेत्रफळ काढण्यासाठी, आपल्याला दिलेल्या राशीला त्याच राशीच्या गुणावे लागेल. <br>"
											+ "म्हणून या जमिनीचे क्षेत्रफळ <br>$= (" + hidecoepow(a, ch, p1)
											+ plusign(b) + hidecoepow(b, ch, p2) + plusign(c) + c + ")\\times("
											+ hidecoepow(a, ch, p1) + plusign(b) + hidecoepow(b, ch, p2) + plusign(c)
											+ c + ")$<br>$= " + m1 + m2 + m3 + "$<br>$= " + st3 + "$<br>$= " + Eqsq
											+ ". . . . .(i)$ <br>या क्षेत्राचे  $" + hidecoepow(d, ch, p3) + plusign(e)
											+ hidecoepow(e, ch, p4)
											+ "$ या समान मापाचे तुकडे करण्यासाठी आपल्याला (i) या समिकरणाला  $"
											+ hidecoepow(d, ch, p3) + plusign(e) + hidecoepow(e, ch, p4)
											+ "$ ने भागावे लागेल <br>दिलेल्या उदाहरणातील भाज्य व भाजकांच्या पदांची मांडणी खालीलप्रमाणे करावी. <br><br>$"
											+ "\\begin{array}{r}" + no_of_parts
											+ "\\phantom{0000000000000} \\\\\\phantom{00000000} "
											+ hidecoepow(d, ch, p3) + plusign(e) + hidecoepow(e, ch, p4)
											+ "\\big)\\overline{" + Eqsq + "} \\\\" + sub1
											+ "\\phantom{000000000000000} \\\\\\underline{-\\phantom{000} -\\phantom{0000000000000000000} }\\\\"
											+ ans1 + "\\phantom{0} \\\\" + sub2
											+ "\\phantom{0000000000}\\\\\\underline{-\\phantom{000} -\\phantom{00000000000000000}} \\\\ "
											+ ans2 + "\\phantom{} \\\\" + sub3
											+ "\\phantom{000}\\\\\\underline{-\\phantom{000} -\\phantom{0000000} } \\\\"
											+ hidecoepow(ans3v, ch, m2p3) + plusign(m3cc) + m3cc
											+ "\\phantom{} \\end{array}$<br><br>"
											+ "हा भागाकार करून आपल्याला भागाकार  $=" + hidecoepow(s1, ch, ex1)
											+ plusign(s2) + hidecoepow(s2, ch, ex2) + plusign(s3)
											+ hidecoepow(s3, ch, ex3) + "$ आणि बाकी $= " + hidecoepow(ans3v, ch, m2p3)
											+ plusign(m3cc) + m3cc
											+ "$<br>$\\Rightarrow$ जमिनीचे सारख्या मापाचे तुकडे   $"
											+ hidecoepow(s1, ch, ex1) + plusign(s2) + hidecoepow(s2, ch, ex2)
											+ plusign(s3) + hidecoepow(s3, ch, ex3) + "$  आणि शिल्लक जमीन  $"
											+ hidecoepow(ans3v, ch, m2p3) + plusign(m3cc) + m3cc
											+ "$ चौ. एकक हे उत्तर .<br>";
									String Solution = soleng + solmar;
									XSSFRow row = sheet1.createRow(i);
									row.createCell(0).setCellValue(i);
									row.createCell(1).setCellValue("Text");
									row.createCell(2).setCellValue(1);
									row.createCell(3).setCellValue("304050404");
									row.createCell(4).setCellValue(Question);
									row.createCell(5).setCellValue(sol);
//                 		row.createCell(6).setCellValue(" ");
//                 		row.createCell(7).setCellValue(" ");
//                 		row.createCell(8).setCellValue(" ");
									row.createCell(9).setCellValue(wrong_ans1);
									row.createCell(10).setCellValue(wrong_ans2);
									row.createCell(11).setCellValue(wrong_ans3);
									row.createCell(12).setCellValue(90);
									row.createCell(13).setCellValue(4);
//                  	row.createCell(14).setCellValue(" ");
									row.createCell(15).setCellValue("siddhimanjarekar274@gmail.com");
									row.createCell(16).setCellValue(Solution);
//          			row.createCell(17).setCellValue(" ");
									row.createCell(18).setCellValue(109);

									mapsize = map.size();
									map.put(Question, i);
									mapsizeafter = map.size();

									// In Java, a map can consist of virtually any number of key-value pairs, but
									// the keys must always be unique — non-repeating.

									if (mapsize == mapsizeafter) {
										System.out.println("duplicate Question" + i + ". " + Question);
										i--;
									}

									if (sol == wrong_ans1 || sol == wrong_ans2 || sol == wrong_ans3
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
			} else {
				continue;
			}
			i++;

		} while (i <= q);
		int rowTotal = sheet1.getLastRowNum();
		// System.out.println(rowTotal);
		XSSFRow row = sheet1.createRow((short) rowTotal + 1);
		row.createCell(0).setCellValue("****");

		// Writing data to the file
		FileOutputStream fileout = new FileOutputStream(filename);
		workbook.write(fileout);
		fileout.close();

		System.out.println("file created");
	}

	public static String plusign(int a) {
		if (a >= 0) {
			return "+";
		}
		return "";
	}

	public static String plusign2(int a, int b) {
		if (a > 0 && b > 0) {
			return a + "+" + b;
		} else if (a > 0 && b < 0) {
			return "-" + a + "+" + Math.abs(b);
		} else if (a < 0 && b > 0) {
			return Math.abs(a) + "-" + b;
		} else if (a < 0 && b < 0) {
			return Math.abs(a) + "+" + Math.abs(b);
		}
		return "";
	}

	public static String plusign3(int a, int b, int c) {
		if (a > 0 && b > 0 && c > 0) {
			return a + "+" + b + "+" + c;
		} else if (a > 0 && b > 0 && c < 0) {
			return "-" + a + "-" + b + "+" + Math.abs(c);
		} else if (a > 0 && b < 0 && c < 0) {
			return "-" + a + "+" + Math.abs(b) + "+" + Math.abs(c);
		} else if (a < 0 && b > 0 && c > 0) {
			return Math.abs(a) + "-" + b + "-" + c;
		} else if (a < 0 && b > 0 && c < 0) {
			return Math.abs(a) + "-" + b + "+" + Math.abs(c);
		} else if (a < 0 && b < 0 && c < 0) {
			return Math.abs(a) + "+" + Math.abs(b) + "+" + Math.abs(c);
		}
		return "";
	}

	public static String hidecoepow(int a, String ch, int p) {
		if (a != 1 && a != 0 && p != 1 && p != 0 && a != -1) {
			return a + ch + "^" + p;
		} else if (a == 1 && p == 1) {
			return ch;
		} else if (a == 0 && p == 0) {
			return "";
		} else if (a == 1 && p == 0) {
			return String.valueOf(a);
		} else if (a == 0 && p == 1) {
			return "";
		} else if (a == -1 && p == 1) {
			return "-" + ch;
		} else if (a == -1 && p == 1) {
			return "-" + ch;
		} else if (a == 1) {
			return ch + "^" + p;
		} else if (a == 0) {
			return "";
		} else if (a == -1) {
			return "-" + ch + "^" + p;
		} else if (p == 0) {
			return String.valueOf(a);
		} else if (p == 1) {
			return a + ch;
		} else {
			return "";
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

	public static String common_sign2(int a, int b) {
		if (a > 0 && b > 0)
			return "+";
		else if (a < 0 && b < 0)
			return "-";
		else if (a > 0 && b < 0)
			return "-";
		else if (a < 0 && b > 0)
			return "-";
		else
			return null;
	}

	public static String common_sign3(int a, int b, int c) {
		if ((a > 0 && b > 0 && c > 0))
			return "+";
		else if ((a > 0 && b < 0 && c < 0) || (a > 0 && b > 0 && c < 0) || (a < 0 && b < 0 && c < 0)
				|| (a < 0 && b > 0 && c < 0) || (a < 0 && b > 0 && c > 0))
			return "-";
		else
			return null;
	}

}
