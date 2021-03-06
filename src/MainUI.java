import javax.swing.*;
import javax.swing.border.LineBorder;
import java.awt.*;
import java.awt.event.KeyEvent;
import java.awt.event.KeyListener;
import java.io.PrintStream;
import java.text.SimpleDateFormat;

/**
 * Created by Shaw on 2015/9/22.
 */
public class MainUI extends JFrame {
	private JTextArea jTA;
	private JLabel jL1, jL2, jL3, jL4;
	private JTextField jTF1;
	private JButton jB1, jB2, jB3;
	private JPanel jP, upperJP;
	private PrintStream psOut, psErr;

	public MainUI() {
		setVisible(true);
		setSize(300, 400);
		setResizable(false);
		setLocationRelativeTo(null);
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setLayout(new BorderLayout());
		addJPanel();
//		this.getRootPane().setDefaultButton(jB1);

		psOut = new PrintStream(System.out) {
			@Override
			public void print(String str) {
				jTA.append(str);
			}

			@Override
			public void print(int i) {
				jTA.append(String.valueOf(i));
			}

			@Override
			public void println(String str) {
				jTA.append(str + "\n");
			}

			@Override
			public void println(int i) {
				jTA.append(String.valueOf(i) + "\n");
			}
		};

		psErr = new PrintStream(System.err) {
			@Override
			public void print(String str) {
				jTA.append(str);
			}

			@Override
			public void print(int i) {
				jTA.append(String.valueOf(i));
			}

			@Override
			public void println(String str) {
				jTA.append(str + "\n");
			}

			@Override
			public void println(int i) {
				jTA.append(String.valueOf(i) + "\n");
			}
		};

		addJTextField();

	}

	public PrintStream getPsOut() {
		return psOut;
	}

	public PrintStream getPsErr() {
		return psErr;
	}

	public void addJPanel() {
		jP = new JPanel();
		add(jP);
		jP.setLayout(null);
		addJButton();
		addJTextArea();
	}

	public void addJButton() {
		jB1 = new JButton("订单");
		jB1.setBounds(30, 300, 100, 20);
		jP.add(jB1);
		jB1.addActionListener(event -> buttonAP("OrderQuerySQL"));

		jB2 = new JButton("提现");
		jB2.setBounds(30, 320, 100, 20);
		jP.add(jB2);
		jB2.addActionListener(event -> buttonAP("WithdrawQuerySQL"));

		jB3 = new JButton("用户");
		jB3.setBounds(30, 340, 100, 20);
		jP.add(jB3);
		jB3.addActionListener(event -> buttonAP("AuthenticatedUserQuerySQL"));
	}

	public void buttonAP(String sql) {
		SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMdd");
		dateFormat.setLenient(false);
		if (jTF1.getText().length() == 8) {
			try {
				dateFormat.parse(jTF1.getText());
				System.out.println("The Date is: " + getDate());
				XLS.makeXLS("PROD", sql, getDate());
			} catch (Exception pe) {
				System.out.println("Wrong Date!");
			}
		} else {
			System.out.println("Wrong Date format!");
		}
	}

	public void addJTextArea() {
		upperJP = new JPanel();
		jP.add(upperJP);
		upperJP.setBounds(5, 5, 285, 230);
		upperJP.setLayout(new GridLayout(1, 1));

		jTA = new JTextArea();
		jTA.setEditable(false);
//		jTA.setLineWrap(true);
		jTA.setBorder(new LineBorder(new Color(127, 157, 185), 1, false));
//		upperJP.add(jTA);
		JScrollPane jSP = new JScrollPane(jTA);
		upperJP.add(jSP);
	}

	public void jTFKL() {
		if (jTF1.getText().length() > 8) {
			jTF1.setText(jTF1.getText().substring(0, 8));
		}
		for (int i = 0; i < jTF1.getText().length(); i++) {
			if (jTF1.getText().charAt(i) < '0' || jTF1.getText().charAt(i) > '9') {
				jTF1.setText("");
				break;
			}
		}
	}

	public void addJTextField() {
		jTF1 = new JTextField("", 8);
		jTF1.setFont(new Font("Consola", Font.BOLD, 20));
		jTF1.addKeyListener(new KeyListener() {
			@Override
			public void keyTyped(KeyEvent e) {
				char keyCh = e.getKeyChar();
				if (keyCh < '0' || keyCh > '9') {
					e.consume();
				}
			}

			@Override
			public void keyPressed(KeyEvent e) {
				jTFKL();
			}

			@Override
			public void keyReleased(KeyEvent e) {
				jTFKL();
			}
		});


		JPanel innerJP1 = new JPanel();
		innerJP1.setBounds(30, 250, 100, 50);
		innerJP1.setLayout(new BorderLayout());
		jP.add(innerJP1);
		innerJP1.add(jTF1, "Center");
		jL1 = new JLabel("日期");
		jL1.setFont(new Font("Consola", Font.BOLD, 15));
		jL1.setHorizontalAlignment(SwingConstants.CENTER);
		innerJP1.add(jL1, "North");

		JPanel innerJP2 = new JPanel();
		innerJP2.setBounds(150, 270, 150, 70);
		innerJP2.setLayout(new BorderLayout());
		jP.add(innerJP2);

		jL2 = new JLabel("日期为八位数字");
		jL2.setFont(new Font("楷体", Font.PLAIN, 15));
		innerJP2.add(jL2, "North");

		jL3 = new JLabel("如20151225");
		jL3.setFont(new Font("楷体", Font.PLAIN, 15));
		innerJP2.add(jL3, "Center");

		jL4 = new JLabel("在当前目录生成报表");
		jL4.setFont(new Font("楷体", Font.PLAIN, 15));
		innerJP2.add(jL4, "South");

	}

	public String getDate() {
		String temp = jTF1.getText();
		return temp.substring(0, 4) + "-" + temp.substring(4, 6) + "-" + temp.substring(6, 8);
	}

	public static void main(String[] args) {
		SwingUtilities.invokeLater(() -> {
					MainUI mainUI = new MainUI();
					System.setOut(mainUI.getPsOut());
					System.setErr(mainUI.getPsErr());
					System.out.println("Application started...");
				}
		);
	}

}
