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

public class Assign5_0304060402_109 {

	public static Scanner scanner = new Scanner(System.in);
	public static Random random = new Random();

	public static void main(String[] args) throws IOException, FileNotFoundException {
		String filename = "C:\\Siddhi\\Maths_Excelsheet\\VML_0304060402_109_Assign5_Siddhi.xlsx";
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
			String girleng[] = { "Smita", "Radhika", "Priya", "Swapnali", "Swara", "Riya", "Sonia", "Jiya", "Jasmine",
					"Rita", "Siya", "Rukhsana", "Muskan" };
			String girlmar[] = { "स्मिता", "राधिका", "प्रिया", "स्वप्नाली", "स्वरा", "रिया", "सोनिया", "जिया",
					"जास्मिन", "रिटा", "सिया", "रुखसाना", "मुस्कान" };
			String boyeng[] = { "Ajay", "Rakesh", "Prem", "Swapnil", "Swaraj", "Amar", "Akbar", "Anthony", "Raheem",
					"Shahrukh", "Salman", "Amir", "Javed" };
			String boymar[] = { "अजय", "राकेश", "प्रेम", "स्वप्नील", "स्वराज", "अमर", "अकबर", "अँथनी", "रहीम", "शाहरुख",
					"सलमान", "अमीर", "जावेद" };
			int p = random.nextInt(girleng.length);
			String name = " ";
			String name1 = " ";
			String gender = " ";
			String gender1 = " ";
			String gender2 = " ";
			String gender3 = " ";
			String gender4 = " ";
			int g = random.nextInt(2);
			if (g == 0) {
				name = girleng[p];
				name1 = girlmar[p];
				gender = "she";
				gender1 = "तिने";
				gender2 = "तिला";
				gender3 = "तिचे";
				gender4 = "तिची";
			} else {
				name = boyeng[p];
				name1 = boymar[p];
				gender = "he";
				gender1 = "त्याने";
				gender2 = "त्याला";
				gender3 = "त्याचे";
				gender4 = "त्याची";
			}
			// Generating random values
			int minA = 50, maxA = 100, minB = 50, maxB = 200;
			int attempted = random.nextInt(maxA - minA + 1) + minA;
			int totalmarks = random.nextInt(maxB - minB + 1) + minB;
			int totalQuestions = 100;
			int correctMarks = 2;
			int incorrectMarks = -1;

			int totalCorrectMarks = attempted * correctMarks;
			int marksDeducted = totalCorrectMarks - totalmarks;
			int incorrectAnswers = marksDeducted / (correctMarks - incorrectMarks);
			if ((marksDeducted % (correctMarks - incorrectMarks)) == 0 && incorrectAnswers > 0) {
				String questionText = "In a competitive examination, there were $100$ questions. "
						+ "The correct answer would carry $2$ marks, and for incorrect answers $1$ mark would be subtracted, and $0$ mark for not attempted questions. "
						+ name + " had attempted $" + attempted + "$ questions and " + gender + " got a total $"
						+ totalmarks + "$ marks. Then how many questions did " + gender
						+ " get wrong?<br>#एका स्पर्धा परीक्षेत $100$ प्रश्न होते. बरोबर उत्तराला $2$ गुण मिळतात, तर प्रत्येक चूक उत्तरासाठी $1$ गुण वजा होतो, आणि न सोडविलेल्या प्रश्नाला $0$ गुण दिले जातात. "
						+ name1 + "ने त्या परीक्षेत $" + attempted + "$ प्रश्न सोडविले आणि " + gender2 + " एकूण $"
						+ totalmarks + "$ गुण मिळाले. तर " + gender1 + " किती प्रश्न चूक सोडविले?<br>";
				String correct_ans = "$" + incorrectAnswers + "$<br>#$" + incorrectAnswers + "$<br>";
				String wrong_ans1 = "$" + (incorrectAnswers + 1) + "$<br>#$" + (incorrectAnswers + 1) + "$<br>";
				String wrong_ans2 = "$" + (incorrectAnswers + 2) + "$<br>#$" + (incorrectAnswers + 2) + "$<br>";
				String wrong_ans3 = "$" + (incorrectAnswers - 1) + "$<br>#$" + (incorrectAnswers - 1) + "$<br>";

				String solution = "Ans : No. of wrong answers $" + incorrectAnswers + "$.<br> Let us assume that "
						+ name
						+ " got $x$ questions correct and $y$ questions wrong.<br> From the given information, we know that "
						+ gender + " has attempted total $" + attempted + "$ questions.<br> $\\therefore x+y="
						+ attempted
						+ ". . . . (i)$<br> Since answers to $x$ questions are correct<br> $\\therefore$ marks for $x$ correct questions $=2x$<br> Also, since $y$ questions were answered wrong<br> $\\therefore$ marks deducted are $y.$<br> $\\therefore$ total marks obtained $=2x-y$<br> But, as given, total marks obtained $"
						+ totalmarks + "$<br> $\\therefore 2x-y=" + totalmarks
						+ " . . . . . (ii)$<br> From equation $(i)$ we get $x=" + attempted
						+ "-y$<br> Substituting this value of $x$ in equation $(ii)$, we get <br> $2(" + attempted
						+ "-y)-y=" + totalmarks + "$<br> $\\therefore " + totalCorrectMarks + "-2y - y= " + totalmarks
						+ "$<br> By solving this equation for $y$, we get<br> $y = " + incorrectAnswers
						+ "$<br>$\\therefore$ no. of wrong answers given by " + name + " are $" + incorrectAnswers
						+ "$ is the answer.<br>#उत्तर : चुकीची दिलेली उत्तरे $" + incorrectAnswers
						+ "$<br> आपण असे मानू की " + name1 + "ने  $x$  प्रश्नांची उत्तरे बरोबर दिली आणि " + gender4
						+ "  $y$  प्रश्नांची उत्तरे चुकली.<br> दिलेल्या माहितीनुसार " + gender1 + " एकूण $" + attempted
						+ "$ प्रश्नांची उत्तरे  दिली. <br> $\\therefore x+y=" + attempted + ". . . . (i)$<br> "
						+ gender1 + " $x$ प्रश्नांची उत्तरे बरोबर दिली आहेत <br> $\\therefore$ " + gender2
						+ " $x$ बरोबर उत्तरांसाठी मिळालेले गुण $=2x$<br> तसेच " + gender1
						+ " $y$ प्रश्नांची उत्तरे चुकीची दिली. <br> $\\therefore$ " + gender3
						+ " $y$ चूक उत्तरांसाठी कमी झालेले गुण $=y.$<br> $\\therefore$ " + gender2
						+ " मिळालेले एकूण गुण $=2x-y$<br> परंतु दिल्यानुसार " + gender2 + " मिळालेले गुण $" + totalmarks
						+ "$ आहेत. <br> $\\therefore 2x-y=" + totalmarks
						+ " . . . . . (ii)$<br> समीकरण $(i)$ नुसार आपल्याला $x=" + attempted
						+ "-y$ असे मिळते. <br> $x$ ची ही किंमत समीकरण $(ii)$ मध्ये ठेवून आपल्याला ,<br> $2(" + attempted
						+ "-y)-y=" + totalmarks + "$ असे मिळते. <br> $\\therefore " + totalCorrectMarks + "-2y - y= " + totalmarks
						+ "$<br> हे समीकरण $y$ साठी सोडविले असता आपल्याला, $y = " + incorrectAnswers
						+ "$ मिळते . <br> $\\therefore$ " + name1 + "ने दिलेली चुकीची उत्तरे $" + incorrectAnswers
						+ "$ आहेत, हे उत्तर. <br>";

				// Create row
				XSSFRow row = sheet1.createRow(i);
				row.createCell(0).setCellValue(i);
				row.createCell(1).setCellValue("Text");
				row.createCell(2).setCellValue(1);
				row.createCell(3).setCellValue("0304060402");
				row.createCell(4).setCellValue(questionText);
				row.createCell(5).setCellValue(correct_ans);
				row.createCell(9).setCellValue(wrong_ans1);
				row.createCell(10).setCellValue(wrong_ans2);
				row.createCell(11).setCellValue(wrong_ans3);
				row.createCell(12).setCellValue(120);
				row.createCell(13).setCellValue(2);
				row.createCell(15).setCellValue("siddhimanjarekar274@gmail.com");
				row.createCell(16).setCellValue(solution);
				row.createCell(18).setCellValue(109);

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
		// System.out.println(rowTotal);
		XSSFRow row = sheet1.createRow((short) rowTotal + 1);
		row.createCell(0).setCellValue("");
		// Writing data to the file
		FileOutputStream fileout = new FileOutputStream(filename);
		workbook.write(fileout);
		fileout.close();

		System.out.println("File created.");

	}

}
