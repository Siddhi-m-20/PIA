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

public class Assign22_0304060401_101 {
	public static Scanner sc = new Scanner(System.in);
	public static Random random = new Random();

	public static void main(String args[]) throws IOException, FileNotFoundException {
		String filename = "C:\\Siddhi\\Maths_Excelsheet\\VML_0304060401_101_Assign22_Siddhi.xlsx";
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Instruction");

		XSSFSheet sheet1 = workbook.createSheet("Questions");

		// Adding header to the Questions sheet
		String[] header = { "Sr. No", "Question Type", "Answer Type", "Topic Number", "Question (Text Only)",
				"Correct Answer 1", "Correct Answer 2", "Correct Answer 3", "Correct Answer 4", "Wrong Answer 1",
				"Wrong Answer 2", "Wrong Answer 3", "Time in seconds", "Difficulty Level",
//				"Question (Image/ Audio/ Video)", "Contributor's Registered mailId", "Solution (Text Only)",
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
			int min = -20;
			int max = 20;
			int a = random.nextInt(max - min + 1) + min;
			a = checkIfZero(a);
			int b = random.nextInt(max - min + 1) + min;
			b = checkIfZero(b);
			int g = random.nextInt(2);
			int c, d;
			if (g == 0) {
				c = -a;
				d = random.nextInt(max - min + 1) + min;
				d = checkIfZero(d);
			} else {
				d = -b;
				c = random.nextInt(max - min + 1) + min;
				c = checkIfZero(c);
			}
			int e = random.nextInt(max - min + 1) + min;
			e = checkIfZero(e);
			int f = random.nextInt(max - min + 1) + min;
			f = checkIfZero(f);
			String cht[] = { "u", "v", "x", "y", "a", "b" };
			int i0 = random.nextInt(cht.length);
			String ch1 = "";
			String ch2 = "";
			if (i0 == 0 || i0 == 1) {
				ch1 = cht[0];
				ch2 = cht[1];
			} else if (i0 == 2 || i0 == 3) {
				ch1 = cht[2];
				ch2 = cht[3];
			} else if (i0 == 4 || i0 == 5) {
				ch1 = cht[4];
				ch2 = cht[5];
			}
			int ac = 0, bd = 0, ef = 0, x, y;
			if (a == -c || b == -d) {
				ac = a + c;
				bd = b + d;
				ef = e + f;
			}
			if (bd != 0 && (ef % bd) == 0 && checkIfNoCommonFactor(a, b) && checkIfNoCommonFactor(c, d)) {
				y = ef / bd;
				if (y != 0) {
					if ((e - (b * y)) % a == 0) {
						x = (e - (b * y)) / a;
						if (x != 0) {
							String Question = "Solve following simultaneous linear equations by elimination of one variable method,<br>$"
									+ hide1a(a, ch1) + ifn(b) + hide1a(b, ch2) + " = " + e + "$ and $" + hide1a(c, ch1)
									+ ifn(d) + hide1a(d, ch2) + " = " + f
									+ "$. <br>#खाली दिलेली एकसामायिक रेषीय समीकरणे एक चल लोप पद्धतीने सोडवा,<br>$"
									+ hide1a(a, ch1) + ifn(b) + hide1a(b, ch2) + " = " + e + "$ आणि $" + hide1a(c, ch1)
									+ ifn(d) + hide1a(d, ch2) + " = " + f + "$. <br>";
							String Correct_ans = "$" + ch1 + " =" + x + "$ and $" + ch2 + " =" + y + "$<br>#$" + ch1
									+ " =" + x + "$ आणि $" + ch2 + " =" + y + "$";
							String wrong_ans1 = "$" + ch1 + " =" + 2 * x + "$ and $" + ch2 + " =" + 2 * y + "$<br>#$"
									+ ch1 + " =" + 2 * x + "$ आणि $" + ch2 + " =" + 2 * y + "$<br>";
							String wrong_ans2 = "$" + ch1 + " =" + (x + 2) + "$ and $" + ch2 + " =" + (y + 3)
									+ "$<br>#$" + ch1 + " =" + (x + 2) + "$ आणि $" + ch2 + " =" + (y + 3) + "$<br>";
							String wrong_ans3 = "$" + ch1 + " =" + a + "$ and $" + ch2 + " =" + b + "$<br>#$" + ch1
									+ "=" + a + "$ आणि $" + ch2 + " =" + b + "$<br>";
							String wrong_ans4 = "$" + ch1 + " =" + c + "$ and $" + ch2 + " =" + d + "$<br>#$" + ch1
									+ "=" + c + "$ आणि $" + ch2 + " =" + d + "$<br>";
							String wrong_ans5 = "$" + ch1 + " =" + ac + "$ and $" + ch2 + " =" + bd + "$<br>#$" + ch1
									+ " =" + ac + "$ आणि $" + ch2 + " =" + bd + "$<br>";
							String wrong_ans6 = "$" + ch1 + " =" + b + "$ and $" + ch2 + " =" + c + "$<br>#$" + ch1
									+ " =" + b + "$ आणि $" + ch2 + " =" + c + "$<br>";
							String Solution = "";
							if (b == -d) {
								Solution = "​Ans : $" + ch1 + " = " + x + "$ and $" + ch2 + " = " + y
										+ "$<br>In elimination of one variable method, we are required to obtain a third (new) equation which will have only one variable by elimination of other variable from the given two equation.<br> To obtain this, we will add given two equations ( add left side of first equation with left side of second equation and similarly for right side) to get<br> $\\begin{alignat*}{13}{}&"
										+ hide1a(a, ch1) + "&{}" + ifn(b) + hide1a(b, ch2) + "{}&={}&" + e
										+ "&{}\\\\+\\quad&{}" + hide1a(c, ch1) + "{}&" + ifn(d) + hide1a(d, ch2)
										+ "&{}={}&" + f + "&{}\\\\\\hline&{}" + hide1a(ac, ch1) + "{}&+0&{}=&{}" + ef
										+ "{}& \\ \\ \\end{alignat*}$<br>Thus $" + ch2
										+ "$ is eliminated here.<br>$\\therefore " + hide1a(ac, ch1) + "=" + ef
										+ "$<br>$\\therefore " + ch1 + " =" + x
										+ "$<br >By substituting this value of $" + ch1
										+ "$ in any of the given equation, we get $" + ch2 + " =" + y
										+ "$<br>$\\therefore " + ch1 + " =" + x + "$ and $" + ch2 + " =" + y
										+ "$ is the answer.<br>#उत्तर :  $" + ch1 + " = " + x + "$ आणि $" + ch2 + " = "
										+ y
										+ "$<br>एकचल लोप पद्धतीत दिलेल्या दोन्ही समीकरणांचा वापर करून, त्यातील एका चलाचा लोप करायचा असतो (एक चल कमी होईल असे करून) जेणेकरून एक नवीन तिसरे समीकरण तयार होईल, आणि त्यात एकच चल पद असेल. <br>या साठी येथे आपण दिलेल्या दोन्ही समीकरणांची बेरीज केली (म्हणजे डाव्या बाजूत डावी बाजू मिळवली आणि उजव्या बाजूत उजवी बाजू  मिळवली ) तर आपल्याला <br > $\\begin{alignat*}{13}{}&"
										+ hide1a(a, ch1) + "&{}" + ifn(b) + hide1a(b, ch2) + "{}&={}&" + e + "&{}"
										+ "\\\\+\\quad&{}" + hide1a(c, ch1) + "{}&" + ifn(d) + hide1a(d, ch2)
										+ "&{}={}&" + f + "&{}" + "\\\\\\hline&{}" + hide1a(ac, ch1) + "{}&+0&{}=&{}"
										+ ef + "{}& \\ \\ \\end{alignat*}$<br>असे मिळते<br> अशा रीतीने आपण $" + ch2
										+ "$ या चलाचा लोप केला आहे. <br>$\\therefore " + hide1a(ac, ch1) + "=" + ef
										+ "$<br>$\\therefore " + ch1 + " =" + x + "$<br >ही मिळालेली $" + ch1
										+ "$ ची किंमत दिलेल्या कोणत्याही एका समीकरणात घालून आपल्याला  $" + ch2 + " ="
										+ y + "$ मिळते. <br >$\\therefore " + ch1 + " =" + x + "$ आणि $" + ch2 + " ="
										+ y + "$ हे उत्तर. <br>";
							} else if (a == -c) {
								Solution = "​Ans : $" + ch1 + " = " + x + "$ and $" + ch2 + " = " + y
										+ "$<br>In elimination of one variable method, we are required to obtain a third (new) equation which will have only one variable by elimination of other variable from the given two equation. <br>To obtain this, we will add given two equations ( add left side of first equation with left side of second equation and similarly for right side) to get<br> $\\begin{alignat*}{13}{}&"
										+ hide1a(a, ch1) + "&{}" + ifn(b) + hide1a(b, ch2) + "{}&={}&" + e
										+ "&{}\\\\+\\quad&{}" + hide1a(c, ch1) + "{}&" + ifn(d) + hide1a(d, ch2)
										+ "&{}={}&" + f + "&{}\\\\\\hline&{}0{}&" + ifn(bd) + hide1a(bd, ch2)
										+ "&{}=&{}" + ef + "{}& \\ \\ \\end{alignat*}$<br>Thus $" + ch1
										+ "$ is eliminated here.<br>" + hideans(bd, ef, y, ch2)
										+ "By substituting this value of $" + ch2
										+ "$ in any of the given equation, we get $" + ch1 + " =" + x
										+ "$<br>$\\therefore " + ch1 + " =" + x + "$ and $" + ch2 + " =" + y
										+ "$ is the answer.<br>#उत्तर :  $" + ch1 + " = " + x + "$ आणि $" + ch2 + " = "
										+ y
										+ "$<br>एकचल लोप पद्धतीत दिलेल्या दोन्ही समीकरणांचा वापर करून, त्यातील एका चलाचा लोप करायचा असतो (एक चल कमी होईल असे करून) जेणेकरून एक नवीन तिसरे समीकरण तयार होईल, आणि त्यात एकच चल पद असेल. <br>या साठी येथे आपण दिलेल्या दोन्ही समीकरणांची बेरीज केली (म्हणजे डाव्या बाजूत डावी बाजू मिळवली आणि उजव्या बाजूत उजवी बाजू मिळवली ) तर आपल्याला <br > $\\begin{alignat*}{13}{}&"
										+ hide1a(a, ch1) + "&{}" + ifn(b) + hide1a(b, ch2) + "{}&={}&" + e
										+ "&{}\\\\+\\quad&{}" + hide1a(c, ch1) + "{}&" + ifn(d) + hide1a(d, ch2)
										+ "&{}={}&" + f + "&{}\\\\\\hline&{}0{}&" + ifn(bd) + hide1a(bd, ch2)
										+ "&{}=&{}" + ef
										+ "{}& \\ \\ \\end{alignat*}$<br>असे मिळते<br> अशा रीतीने आपण $" + ch1
										+ "$ या चलाचा लोप केला आहे. <br>" + hideans(bd, ef, y, ch2) + "ही मिळालेली $"
										+ ch2 + "$ ची किंमत दिलेल्या कोणत्याही एका समीकरणात घालून आपल्याला  $" + ch1
										+ " =" + x + "$ मिळते. <br >$\\therefore " + ch1 + " =" + x + "$ आणि $" + ch2
										+ " =" + y + "$ हे उत्तर. <br>";
							}
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
							row.createCell(3).setCellValue("0304060401");
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
							row.createCell(13).setCellValue(1); // row.createCell(14).setCellValue(" ");
							row.createCell(15).setCellValue("siddhimanjarekar274@gmail.com");
							row.createCell(16).setCellValue(Solution); //
							row.createCell(17).setCellValue(" ");
							row.createCell(18).setCellValue(101);

							if (Correct_ans == wrong_ans1 || Correct_ans == wrong_ans2 || Correct_ans == wrong_ans3
									|| Correct_ans == wrong_ans4 || Correct_ans == wrong_ans5
									|| Correct_ans == wrong_ans6 || wrong_ans1 == wrong_ans2 || wrong_ans1 == wrong_ans3
									|| wrong_ans1 == wrong_ans4 || wrong_ans1 == wrong_ans5 || wrong_ans1 == wrong_ans6
									|| wrong_ans2 == wrong_ans3 || wrong_ans2 == wrong_ans4 || wrong_ans2 == wrong_ans5
									|| wrong_ans2 == wrong_ans6 || wrong_ans3 == wrong_ans4 || wrong_ans3 == wrong_ans4
									|| wrong_ans3 == wrong_ans6 || wrong_ans4 == wrong_ans5 || wrong_ans4 == wrong_ans6
									|| wrong_ans5 == wrong_ans6) {
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
			} else if (ac != 0 && (ef % ac) == 0 && checkIfNoCommonFactor(a, b) && checkIfNoCommonFactor(c, d)) {
				x = ef / ac;
				if (x != 0) {
					if ((e - (a * x)) % b == 0) {
						y = (e - (a * x)) / b;
						if (y != 0) {
							String Question = "Solve following simultaneous linear equations by elimination of one variable method,<br>$"
									+ hide1a(a, ch1) + ifn(b) + hide1a(b, ch2) + " = " + e + "$ and $" + hide1a(c, ch1)
									+ ifn(d) + hide1a(d, ch2) + " = " + f
									+ "$. <br>#खाली दिलेली एकसामायिक रेषीय समीकरणे एक चल लोप पद्धतीने सोडवा,<br>$"
									+ hide1a(a, ch1) + ifn(b) + hide1a(b, ch2) + " = " + e + "$ आणि $" + hide1a(c, ch1)
									+ ifn(d) + hide1a(d, ch2) + " = " + f + "$. <br>";
							String Correct_ans = "$" + ch1 + " =" + x + "$ and $" + ch2 + " =" + y + "$<br>#$" + ch1
									+ " =" + x + "$ आणि $" + ch2 + " =" + y + "$";
							String wrong_ans1 = "$" + ch1 + " =" + 2 * x + "$ and $" + ch2 + " =" + 2 * y + "$<br>#$"
									+ ch1 + " =" + 2 * x + "$ आणि $" + ch2 + " =" + 2 * y + "$<br>";
							String wrong_ans2 = "$" + ch1 + " =" + (x + 3) + "$ and $" + ch2 + " =" + (y + 3)
									+ "$<br>#$" + ch1 + " =" + (x + 3) + "$ आणि $" + ch2 + " =" + (y + 3) + "$<br>";
							String wrong_ans3 = "$" + ch1 + " =" + a + "$ and $" + ch2 + " =" + c + "$<br>#$" + ch1
									+ " =" + a + "$ आणि $" + ch2 + " =" + c + "$<br>";
							String wrong_ans4 = "$" + ch1 + " =" + (x - 2) + "$ and $" + ch2 + " =" + (y - 3)
									+ "$<br>#$" + ch1 + " =" + (x - 2) + "$ आणि $" + ch2 + " =" + (y - 3) + "$<br>";
							String wrong_ans5 = "$" + ch1 + " =" + b + "$ and $" + ch2 + " =" + d + "$<br>#$" + ch1
									+ " =" + b + "$ आणि $" + ch2 + " =" + d + "$<br>";
							String wrong_ans6 = "$" + ch1 + " =" + bd + "$ and $" + ch2 + " =" + ac + "$<br>#$" + ch1
									+ " =" + bd + "$ आणि $" + ch2 + " =" + ac + "$<br>";
							String Solution = "";
							if (b == -d) {
								Solution = "​Ans : $" + ch1 + " = " + x + "$ and $" + ch2 + " = " + y
										+ "$<br>In elimination of one variable method, we are required to obtain a third (new) equation which will have only one variable by elimination of other variable from the given two equation.<br> To obtain this, we will add given two equations ( add left side of first equation with left side of second equation and similarly for right side) to get<br> $\\begin{alignat*}{13}{}&"
										+ hide1a(a, ch1) + "&{}" + ifn(b) + hide1a(b, ch2) + "{}&={}&" + e
										+ "&{}\\\\+\\quad&{}" + hide1a(c, ch1) + "{}&" + ifn(d) + hide1a(d, ch2)
										+ "&{}={}&" + f + "&{}\\\\\\hline&{}" + hide1a(ac, ch1) + "{}&+0&{}=&{}" + ef
										+ "{}& \\ \\ \\end{alignat*}$<br>Thus $" + ch2
										+ "$ is eliminated here.<br>$\\therefore " + hide1a(ac, ch1) + "=" + ef
										+ "$<br>$\\therefore " + ch1 + " =" + x
										+ "$<br >By substituting this value of $" + ch1
										+ "$ in any of the given equation, we get $" + ch2 + " =" + y
										+ "$<br>$\\therefore " + ch1 + " =" + x + "$ and $" + ch2 + " =" + y
										+ "$ is the answer.<br>#उत्तर :  $" + ch1 + " = " + x + "$ आणि $" + ch2 + " = "
										+ y
										+ "$<br>एकचल लोप पद्धतीत दिलेल्या दोन्ही समीकरणांचा वापर करून, त्यातील एका चलाचा लोप करायचा असतो (एक चल कमी होईल असे करून) जेणेकरून एक नवीन तिसरे समीकरण तयार होईल, आणि त्यात एकच चल पद असेल. <br>या साठी येथे आपण दिलेल्या दोन्ही समीकरणांची बेरीज केली (म्हणजे डाव्या बाजूत डावी बाजू मिळवली आणि उजव्या बाजूत उजवी बाजू  मिळवली ) तर आपल्याला <br > $\\begin{alignat*}{13}{}&"
										+ hide1a(a, ch1) + "&{}" + ifn(b) + hide1a(b, ch2) + "{}&={}&" + e + "&{}"
										+ "\\\\+\\quad&{}" + hide1a(c, ch1) + "{}&" + ifn(d) + hide1a(d, ch2)
										+ "&{}={}&" + f + "&{}" + "\\\\\\hline&{}" + hide1a(ac, ch1) + "{}&+0&{}=&{}"
										+ ef + "{}& \\ \\ \\end{alignat*}$<br>असे मिळते<br> अशा रीतीने आपण $" + ch2
										+ "$ या चलाचा लोप केला आहे. <br>$\\therefore " + hide1a(ac, ch1) + "=" + ef
										+ "$<br>$\\therefore " + ch1 + " =" + x + "$<br >ही मिळालेली $" + ch1
										+ "$ ची किंमत दिलेल्या कोणत्याही एका समीकरणात घालून आपल्याला  $" + ch2 + " ="
										+ y + "$ मिळते. <br >$\\therefore " + ch1 + " =" + x + "$ आणि $" + ch2 + " ="
										+ y + "$ हे उत्तर. <br>";
							} else if (a == -c) {
								Solution = "​Ans : $" + ch1 + " = " + x + "$ and $" + ch2 + " = " + y
										+ "$<br>In elimination of one variable method, we are required to obtain a third (new) equation which will have only one variable by elimination of other variable from the given two equation. <br>To obtain this, we will add given two equations ( add left side of first equation with left side of second equation and similarly for right side) to get<br> $\\begin{alignat*}{13}{}&"
										+ hide1a(a, ch1) + "&{}" + ifn(b) + hide1a(b, ch2) + "{}&={}&" + e
										+ "&{}\\\\+\\quad&{}" + hide1a(c, ch1) + "{}&" + ifn(d) + hide1a(d, ch2)
										+ "&{}={}&" + f + "&{}\\\\\\hline&{}0{}&" + ifn(bd) + hide1a(bd, ch2)
										+ "&{}=&{}" + ef + "{}& \\ \\ \\end{alignat*}$<br>Thus $" + ch1
										+ "$ is eliminated here.<br>" + hideans(bd, ef, y, ch2)
										+ "By substituting this value of $" + ch2
										+ "$ in any of the given equation, we get $" + ch1 + " =" + x
										+ "$<br>$\\therefore " + ch1 + " =" + x + "$ and $" + ch2 + " =" + y
										+ "$ is the answer.<br>#उत्तर :  $" + ch1 + " = " + x + "$ आणि $" + ch2 + " = "
										+ y
										+ "$<br>एकचल लोप पद्धतीत दिलेल्या दोन्ही समीकरणांचा वापर करून, त्यातील एका चलाचा लोप करायचा असतो (एक चल कमी होईल असे करून) जेणेकरून एक नवीन तिसरे समीकरण तयार होईल, आणि त्यात एकच चल पद असेल. <br>या साठी येथे आपण दिलेल्या दोन्ही समीकरणांची बेरीज केली (म्हणजे डाव्या बाजूत डावी बाजू मिळवली आणि उजव्या बाजूत उजवी बाजू मिळवली ) तर आपल्याला <br > $\\begin{alignat*}{13}{}&"
										+ hide1a(a, ch1) + "&{}" + ifn(b) + hide1a(b, ch2) + "{}&={}&" + e
										+ "&{}\\\\+\\quad&{}" + hide1a(c, ch1) + "{}&" + ifn(d) + hide1a(d, ch2)
										+ "&{}={}&" + f + "&{}\\\\\\hline&{}0{}&" + ifn(bd) + hide1a(bd, ch2)
										+ "&{}=&{}" + ef
										+ "{}& \\ \\ \\end{alignat*}$<br>असे मिळते<br> अशा रीतीने आपण $" + ch1
										+ "$ या चलाचा लोप केला आहे. <br>" + hideans(bd, ef, y, ch2) + "ही मिळालेली $"
										+ ch2 + "$ ची किंमत दिलेल्या कोणत्याही एका समीकरणात घालून आपल्याला  $" + ch1
										+ " =" + x + "$ मिळते. <br >$\\therefore " + ch1 + " =" + x + "$ आणि $" + ch2
										+ " =" + y + "$ हे उत्तर. <br>";
							}
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
							row.createCell(3).setCellValue("0304060401");
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
							row.createCell(13).setCellValue(1); // row.createCell(14).setCellValue(" ");
							row.createCell(15).setCellValue("siddhimanjarekar274@gmail.com");
							row.createCell(16).setCellValue(Solution); //
							row.createCell(17).setCellValue(" ");
							row.createCell(18).setCellValue(101);

							if (Correct_ans == wrong_ans1 || Correct_ans == wrong_ans2 || Correct_ans == wrong_ans3
									|| Correct_ans == wrong_ans4 || Correct_ans == wrong_ans5
									|| Correct_ans == wrong_ans6 || wrong_ans1 == wrong_ans2 || wrong_ans1 == wrong_ans3
									|| wrong_ans1 == wrong_ans4 || wrong_ans1 == wrong_ans5 || wrong_ans1 == wrong_ans6
									|| wrong_ans2 == wrong_ans3 || wrong_ans2 == wrong_ans4 || wrong_ans2 == wrong_ans5
									|| wrong_ans2 == wrong_ans6 || wrong_ans3 == wrong_ans4 || wrong_ans3 == wrong_ans4
									|| wrong_ans3 == wrong_ans6 || wrong_ans4 == wrong_ans5 || wrong_ans4 == wrong_ans6
									|| wrong_ans5 == wrong_ans6) {
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
			i++;
		} while (i < q + 1);

		int rowTotal = sheet1.getLastRowNum();
//          System.out.println(rowTotal);
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
		else if (a == -1)
			return "-" + chr;
		return null;

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

	public static String hideans(int bd, int ef, int y, String ch2) {
		if (bd == 1) {
			return "$\\therefore " + hide1a(bd, ch2) + "=" + ef + "$<br>";
		} else {
			return "$\\therefore " + hide1a(bd, ch2) + "=" + ef + "$<br>$\\therefore " + ch2 + " =" + y + "$<br >";
		}
	}
}