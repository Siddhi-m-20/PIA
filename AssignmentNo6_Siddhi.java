package com.studentassignment;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Random;
import java.util.Vector;
import java.util.Scanner;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class AssignmentNo6_Siddhi {
	public static Scanner sc = new Scanner(System.in);
	public static Random random = new Random();

	public static void main(String args[]) throws IOException, FileNotFoundException {

		String filename = "E:\\VESIT_Students_Sheets\\ExelSheet\\VML_03040405_115_Assign6_Siddhi.xlsx";
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
			int chr = random.nextInt(cht.length);
			int chr1 = random.nextInt(cht.length);
			if (chr != chr1) {
				int chr2 = random.nextInt(cht.length);
				if (chr != chr2 && chr1 != chr2) {
					String ch1 = cht[chr];
					String ch2 = cht[chr1];
					String ch3 = cht[chr2];
					// powers
					int pow[] = { 1, 2, 3, 1 };
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
								String term1 = if1(a) + power(ch1, p1) + power(ch2, p2);
								String term2 = if1(b) + power(ch1, p1) + power(ch2, p3) + power(ch3, p2);
								String term3 = if1(c) + power(ch2, p4) + power(ch3, p2);
								int ab = a * b;
								int ac = a * c;
								int p1p1 = p1 + p1;
								int p2p3 = p2 + p3;
								int p2p4 = p2 + p4;
								String Question = "Factors of a certain algebraic expression are $" + term1 + "$ and $("
										+ term2 + ifn(c) + term3
										+ ")$. What is that expression.<br># एका बैजिक राशीचे अवयव   $" + term1
										+ "$ आणि $(" + term2 + ifn(c) + term3 + ")$ आहेत तर ती राशी कोणती?<br>";
								// Generate Correct answer
								String Correct_ans = "$" + ab + power(ch1, p1p1) + power(ch2, p2p3) + power(ch3, p2)
										+ ifn(ac) + ac + power(ch1, p1) + power(ch2, p2p4) + power(ch3, p2) + "$<br>";
								// Generate wrong options
								String wrong_ans1 = "$" + (a + b) + power(ch1, p1p1) + power(ch2, p2p3) + power(ch3, p2)
										+ ifn(a + c) + (a + c) + power(ch1, p1) + power(ch2, p2p4) + power(ch3, p2)
										+ "$<br>";
								String wrong_ans2 = "$" + ab + power(ch1, (p1p1 + 1)) + power(ch2, (p2p3 + 2))
										+ power(ch3, (p2 + 2)) + ifn(ac) + ac + power(ch1, p1) + power(ch2, p2p4)
										+ "$<br>";
								String wrong_ans3 = "$" + ab + power(ch1, p1p1) + power(ch2, p2p3) + power(ch3, p2)
										+ ifn(ac) + ac + power(ch1, (p1 + 1)) + power(ch2, p2p4) + power(ch3, p2)
										+ "$<br>";
								String wrong_ans4 = "$" + (a + b) + power(ch1, (p1p1 + 1)) + power(ch2, (p2p3 + 2))
										+ power(ch3, p2) + ifn(a + c) + (a + c) + power(ch1, (p1 + 1))
										+ power(ch3, (p2 + 2)) + "$<br>";
								String wrong_ans5 = "$" + term1 + ifn(b) + term2 + ifn(c) + term3 + "$<br>";
								String Solution = "Ans : ​​​​" + Correct_ans
										+ "We know that factor of an algebraic expression is that expression which divides the given expression with zero remainder.<br> It means, both given expressions $"
										+ term1 + "$ and $(" + term2 + ifn(c) + term3
										+ ")$ are divisors with zero remainder, of the required expression.<br> "
										+ "In short, if we take the product of these two factors we will get the required expression.<br>$\\therefore$ let us find out the product of $("
										+ term1 + ")$ and $(" + term2 + ifn(c) + term3
										+ ")$<br>$\\therefore$ required expression is,<br>$\\therefore (" + term1
										+ ")\\times (" + term2 + ifn(c) + term3 + ")\\ . . . .$ by taking multiplication of $"
										+ term1 + "$ with each term in the bracket.<br>$= " + ab + power(ch1, p1p1)
										+ power(ch2, p2p3) + power(ch3, p2) + ifn(ac) + ac + power(ch1, p1)
										+ power(ch2, p2p4) + power(ch3, p2)
										+ "$ is the desired expression, is the answer.<br>#​उत्तर : ​​​​" + Correct_ans
										+ "​​बैजिक राशीचे अवयव म्हणजे फक्त त्या राशी, ज्यांनी दिलेल्या राशीला निःशेष (बाकी शून्य) भाग जातो.<br> याचा अर्थ आपल्याला हव्या असलेल्या राशीला $"
										+ term1 + "$ आणि $(" + term2 + ifn(c) + term3
										+ ")$ या दोन्ही अवयवांनी निःशेष भाग जायला  हवा.<br> "
										+ "म्हणजेच जर आपण या दोन अवयवांचा गुणाकार केला तर आपल्याला हवी असलेली राशी मिळेल.<br>"
										+ "$\\therefore$ आपण $(" + term1 + ")$ आणि $(" + term2 + ifn(c) + term3
										+ ")$ या दोन राशींचा गुणाकार शोधू. <br>$\\therefore (" + term1 + ")\\times ("
										+ term2 + ifn(c) + term3 + ")\\ . . . . " + term1
										+ "$ या राशीने कंसातील प्रत्येक पदाला गुणून <br>$= " + ab + power(ch1, p1p1)
										+ power(ch2, p2p3) + power(ch3, p2) + ifn(ac) + ac + power(ch1, p1)
										+ power(ch2, p2p4) + power(ch3, p2) + "$ ही हवी असलेली राशी आहे, हे उत्तर.<br>";
								
								
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
								row.createCell(13).setCellValue(3);
//                  			row.createCell(14).setCellValue(" ");
								row.createCell(15).setCellValue("siddhimanjarekar274@gmail.com");
								row.createCell(16).setCellValue(Solution);
//                  			row.createCell(17).setCellValue(" ");
								row.createCell(18).setCellValue(115);
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
										|| wrong_ans2 == wrong_ans3||(a+b)==1||(a+b)==-1||(a+c)==1||(a+c)==-1||ab==0||(a+b)==0||(a+c)==0||ac==0||ab==1||ab==-1||ac==1||ac==-1) {
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
		row1.createCell(0).setCellValue("****");

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
