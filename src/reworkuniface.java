import java.awt.*;
import java.awt.event.*;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import javax.swing.*;
import javax.swing.JFormattedTextField.AbstractFormatter;
import javax.swing.event.ListSelectionEvent;
import javax.swing.event.ListSelectionListener;
import javax.swing.text.DefaultCaret;

import org.jdatepicker.impl.JDatePanelImpl;
import org.jdatepicker.impl.JDatePickerImpl;
import org.jdatepicker.impl.UtilDateModel;
import java.util.*;
import java.util.Calendar;
import java.text.*;

/**
 * Cette classe représente une interface utilisateur pour le programme de restructuration UNIFACE.
 */
public class reworkuniface extends JFrame implements ListSelectionListener {
    FlowLayout experimentLayout = new FlowLayout();
    String datedebut = "";
    String datefin = "";
    int selectedIndex;

    static String getText;

    /**
     * Constructeur de la classe reworkuniface.
     *
     * @param name Le nom de la fenêtre.
     */
    public reworkuniface(String name) {
        super(name);
    }

    /**
     * Ajoute les composants à l'interface graphique.
     *
     * @param pane Le conteneur de l'interface.
     */
    public void addComponentsToPane(final Container pane) {
        final JPanel compsToExperiment = new JPanel();
        compsToExperiment.setLayout(new GridLayout(11, 6));
        JPanel controls = new JPanel();
        controls.setLayout(new FlowLayout());

        //LISTBUTTON

        //BUTTON
        JButton medspecheck, medidlist, uniquedos, healhunit, oneday, review, clear, exit;
        Rectangle rmedspecheck, rmedidlist, runiquedos, rhealhunit, roneday, rreview, rclear, rexit;

        medidlist = new JButton("Code Medecin");
        medspecheck = new JButton("Valider Code Med et Spécialité");
        healhunit = new JButton("Liste Unité de Soin");
        uniquedos = new JButton("Liste Dossier Unique");
        oneday = new JButton("Liste One Day");
        review = new JButton("Prévisualisation");
        clear = new JButton("Remise à Zéro");
        exit = new JButton("Quitter");
        pane.add(compsToExperiment, BorderLayout.NORTH);
        pane.add(controls, BorderLayout.SOUTH);

        rmedidlist = new Rectangle(20, 50, 80, 40);
        rmedspecheck = new Rectangle(20, 50, 80, 40);
        runiquedos = new Rectangle(20, 50, 80, 40);
        rhealhunit = new Rectangle(20, 50, 80, 40);
        roneday = new Rectangle(20, 50, 80, 40);
        rreview = new Rectangle(20, 50, 80, 40);
        rclear = new Rectangle(20, 50, 80, 40);
        rexit = new Rectangle(20, 50, 80, 40);

        medidlist.setBounds(rmedidlist);
        medspecheck.setBounds(rmedspecheck);
        uniquedos.setBounds(runiquedos);
        healhunit.setBounds(rhealhunit);
        oneday.setBounds(roneday);
        review.setBounds(rreview);
        clear.setBounds(rclear);
        exit.setBounds(rexit);

        //BEGINCALENDAR

        /**
         * Cette classe définit un formatteur de date pour le datePicker2.
         */
        class DateLabelFormatter2 extends AbstractFormatter {

            private String datePattern = "yyyMMdd";
            private SimpleDateFormat dateFormatter = new SimpleDateFormat(datePattern);

            @Override
            public Object stringToValue(String text2) throws ParseException {
                return dateFormatter.parseObject(text2);
            }

            @Override
            public String valueToString(Object value2) throws ParseException {
                if (value2 != null) {
                    Calendar cal2 = (Calendar) value2;
                    datedebut = dateFormatter.format(cal2.getTime());
                    return dateFormatter.format(cal2.getTime());
                } else {
                    DateFormat dateFormat = new SimpleDateFormat("yyyyMMdd");
                    Calendar col = Calendar.getInstance();
                    java.util.Date daydate = col.getTime();
                    datedebut = dateFormat.format(daydate);
                }
                return datedebut;
            }

        }

        UtilDateModel model2 = new UtilDateModel();
        Properties p2 = new Properties();
        p2.put("text.today", "Today");
        p2.put("text.month", "Month");
        p2.put("text.year", "Year");
        JDatePanelImpl datePanel2 = new JDatePanelImpl(model2, p2);
        JDatePickerImpl datePicker2 = new JDatePickerImpl(datePanel2, new DateLabelFormatter2());
        setLayout(new GridBagLayout());

        //ENDCALENDAR

        /**
         * Cette classe définit un formatteur de date pour le datePicker.
         */
        class DateLabelFormatter extends AbstractFormatter {

            private String datePattern = "yyyMMdd";
            private SimpleDateFormat dateFormatter = new SimpleDateFormat(datePattern);

            @Override
            public Object stringToValue(String text) throws ParseException {
                return dateFormatter.parseObject(text);
            }

            @Override
            public String valueToString(Object value) throws ParseException {
                if (value != null) {
                    Calendar cal = (Calendar) value;
                    datefin = dateFormatter.format(cal.getTime());
                    return dateFormatter.format(cal.getTime());
                }
                return "";
            }

        }

        UtilDateModel model = new UtilDateModel();
        Properties p = new Properties();
        p.put("text.today", "Today");
        p.put("text.month", "Month");
        p.put("text.year", "Year");
        JDatePanelImpl datePanel = new JDatePanelImpl(model, p);
        JDatePickerImpl datePicker = new JDatePickerImpl(datePanel, new DateLabelFormatter());
        setLayout(new GridBagLayout());

        //SPECIALITE
        Object[] specialites = new Object[]{"Spécialité Non Spécifié", "Chirurgie Generale", "Chirurgie Plastique",
                "Chirurgie Abdominale", "Chirurgie Vasculaire", "Gynecologie", "Ophtalmologie", "O.R.L", "Urologie",
                "Orthopedie", "Stomatologie", "Dermato", "Gastro-Enterologie", "Anesthesie", "ASSISTANCE OPERATOIRE",
                "MEDECIN DE FAMILLE", "URGENCE", "Imagierie medicale", "MATERNITE", "Medecine generale",
                "Medecine generale 2", "Assistant Chirurgie", "Assistant Gynecologie", "Assistant Orthopedie",
                "ANESTHESIE", "Neuro_Chirurgie", "O.R.L + Readaptation", "Medecine interne", "Pneumologie",
                "Pediatrie", "CHIRURGIE MAIN", "assistant anesthesie", "CHIR THORACIQUE"};
        JComboBox spelist = new JComboBox(specialites);

        //CODE MEDECIN

        //1
        JPanel codemed1 = new JPanel();
        JTextField jtf1 = new JTextField("");
        JLabel label1 = new JLabel("MED CODE 1");
        codemed1.setBackground(Color.white);
        codemed1.setLayout(new BorderLayout());
        JPanel top1 = new JPanel();
        Font police1 = new Font("Arial", Font.BOLD, 10);
        jtf1.setFont(police1);
        jtf1.setPreferredSize(new Dimension(45, 20));
        jtf1.setForeground(Color.BLUE);
        top1.add(label1);
        top1.add(jtf1);
        codemed1.add(top1, BorderLayout.NORTH);

        //2
        JPanel codemed2 = new JPanel();
        JTextField jtf2 = new JTextField("");
        JLabel label2 = new JLabel("MED CODE 2");
        codemed2.setBackground(Color.white);
        codemed2.setLayout(new BorderLayout());
        JPanel top2 = new JPanel();
        Font police2 = new Font("Arial", Font.BOLD, 10);
        jtf2.setFont(police2);
        jtf2.setPreferredSize(new Dimension(45, 20));
        jtf2.setForeground(Color.BLUE);
        top2.add(label2);
        top2.add(jtf2);
        codemed2.add(top2, BorderLayout.NORTH);

        //3
        JPanel codemed3 = new JPanel();
        JTextField jtf3 = new JTextField("");
        JLabel label3 = new JLabel("MED CODE 3");
        codemed3.setBackground(Color.white);
        codemed3.setLayout(new BorderLayout());
        JPanel top3 = new JPanel();
        Font police3 = new Font("Arial", Font.BOLD, 10);
        jtf3.setFont(police3);
        jtf3.setPreferredSize(new Dimension(45, 20));
        jtf3.setForeground(Color.BLUE);
        top3.add(label3);
        top3.add(jtf3);
        codemed3.add(top3, BorderLayout.NORTH);

        //4
        JPanel codemed4 = new JPanel();
        JTextField jtf4 = new JTextField("");
        JLabel label4 = new JLabel("MED CODE 4");
        codemed4.setBackground(Color.white);
        codemed4.setLayout(new BorderLayout());
        JPanel top4 = new JPanel();
        Font police4 = new Font("Arial", Font.BOLD, 10);
        jtf4.setFont(police4);
        jtf4.setPreferredSize(new Dimension(45, 20));
        jtf4.setForeground(Color.BLUE);
        top4.add(label4);
        top4.add(jtf4);
        codemed4.add(top4, BorderLayout.NORTH);

        //5
        JPanel codemed5 = new JPanel();
        JTextField jtf5 = new JTextField("");
        JLabel label5 = new JLabel("MED CODE 5");
        codemed5.setBackground(Color.white);
        codemed5.setLayout(new BorderLayout());
        JPanel top5 = new JPanel();
        Font police5 = new Font("Arial", Font.BOLD, 10);
        jtf5.setFont(police5);
        jtf5.setPreferredSize(new Dimension(45, 20));
        jtf5.setForeground(Color.BLUE);
        top5.add(label5);
        top5.add(jtf5);
        codemed5.add(top5, BorderLayout.NORTH);

        //"6"
        JPanel codemed6 = new JPanel();
        JTextField jtf6 = new JTextField("");
        JLabel label6 = new JLabel("                 ");
        codemed6.setLayout(new BorderLayout());
        JPanel top6 = new JPanel();
        Font police6 = new Font("Arial", Font.BOLD, 8);
        jtf6.setFont(police6);
        jtf6.setPreferredSize(new Dimension(0, 0));
        jtf6.setForeground(Color.BLUE);
        top6.add(label6);
        top6.add(jtf6);
        codemed6.add(top6, BorderLayout.NORTH);

        // FENETRE CODE MEDECIN

        // AJOUT CONTENT

        JTextArea loadingarea = new JTextArea(1, 4);
        Font fontloadarea = new Font("Arial", Font.BOLD, 15);
        loadingarea.setFont(fontloadarea);
        loadingarea.setForeground(Color.BLUE);

        compsToExperiment.add(new Label("                                                                                       Spécialités : "));
        compsToExperiment.add(spelist);
        compsToExperiment.add(new Label("                                       Choisissez la date de début de recherche : "));
        compsToExperiment.add(datePicker2);
        compsToExperiment.add(new Label("                                       Choisissez la date de fin de recherche : "));
        compsToExperiment.add(datePicker);
        compsToExperiment.add(codemed1);
        compsToExperiment.add(codemed4);
        compsToExperiment.add(codemed2);
        compsToExperiment.add(codemed5);
        compsToExperiment.add(codemed3);
        compsToExperiment.add(codemed6);
        compsToExperiment.add(medidlist);
        compsToExperiment.add(healhunit);
        compsToExperiment.add(uniquedos);
        compsToExperiment.add(oneday);
        compsToExperiment.add(review);
        compsToExperiment.add(clear);
        compsToExperiment.add(loadingarea);

        //compsToExperiment.add(exit);

        //ACTION BUTTON

        medidlist.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                // open a new frame i.e window
                JFrame medwindow = new JFrame("Code Medecin");
                medwindow.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
                JPanel panel = new JPanel();
                panel.setLayout(new BoxLayout(panel, BoxLayout.Y_AXIS));
                panel.setOpaque(true);
                JTextArea textArea = new JTextArea(15, 50);
                textArea.setFont(Font.getFont(Font.SANS_SERIF));
                JScrollPane scroller = new JScrollPane(textArea);
                scroller.setVerticalScrollBarPolicy(ScrollPaneConstants.VERTICAL_SCROLLBAR_ALWAYS);
                scroller.setHorizontalScrollBarPolicy(ScrollPaneConstants.HORIZONTAL_SCROLLBAR_NEVER);
                JPanel inputpanel = new JPanel();
                inputpanel.setLayout(new FlowLayout());
                JTextField input = new JTextField(20);
                JButton button = new JButton("Rechercher");
                DefaultCaret caret = (DefaultCaret) textArea.getCaret();
                caret.setUpdatePolicy(DefaultCaret.ALWAYS_UPDATE);
                panel.add(scroller);
                inputpanel.add(input);
                inputpanel.add(button);
                panel.add(inputpanel);
                medwindow.getContentPane().add(BorderLayout.CENTER, panel);
                medwindow.pack();
                medwindow.setLocationByPlatform(true);
                medwindow.setVisible(true);
                medwindow.setResizable(false);
                input.requestFocus();

                button.addActionListener(new ActionListener() {
                    public void actionPerformed(ActionEvent e) {
                        textArea.setText("");
                        String valuemedname = input.getText();
                        String newligne = System.getProperty("line.separator");
                        try {
                            Class.forName("");
                            Connection con = DriverManager.getConnection("");
                            String query = "Select MEDNB , MEDNAME FROM MEDBLOC WHERE MEDNAME LIKE '" + valuemedname + "%' ";
                            Statement stmt = con.createStatement();
                            ResultSet rs = stmt.executeQuery(query);
                            for (int i = 0; i < 5; i++) {
                                rs.next();
                                String mednb = rs.getString("MEDNB");
                                String medname = rs.getString("MEDNAME");
                                String sortie = mednb + newligne + medname + newligne + newligne;
                                System.out.println(sortie);
                                textArea.append(sortie);
                            }
                        } catch (SQLException e2) {
                            e2.printStackTrace();
                        } catch (ClassNotFoundException e2) {
                            e2.printStackTrace();
                        } finally {

                        }
                    }
                });
            }
        });

        uniquedos.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                int valuetosend = spelist.getSelectedIndex();
                String valuemed1 = jtf1.getText();
                String valuemed2 = jtf2.getText();
                String valuemed3 = jtf3.getText();
                String valuemed4 = jtf4.getText();
                String valuemed5 = jtf5.getText();

                reworkuniquedos nw = new reworkuniquedos();
                nw.writeuniquedos(datedebut, datefin, valuetosend, valuemed1, valuemed2, valuemed3, valuemed4, valuemed5, loadingarea);
            }
        });

        oneday.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                int valuetosend = spelist.getSelectedIndex();
                String valuemed1 = jtf1.getText();
                String valuemed2 = jtf2.getText();
                String valuemed3 = jtf3.getText();
                String valuemed4 = jtf4.getText();
                String valuemed5 = jtf5.getText();

                reworkoneday nw = new reworkoneday();
                nw.writeoneday(datedebut, datefin, valuetosend, valuemed1, valuemed2, valuemed3, valuemed4, valuemed5, loadingarea);
            }
        });

        uniquedos.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {

            }
        });

        healhunit.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                int valuetosend = spelist.getSelectedIndex();
                String valuemed1 = jtf1.getText();
                String valuemed2 = jtf2.getText();
                String valuemed3 = jtf3.getText();
                String valuemed4 = jtf4.getText();
                String valuemed5 = jtf5.getText();

                reworkhealthunit nw = new reworkhealthunit();
                nw.writehealthunit(datedebut, datefin, valuetosend, valuemed1, valuemed2, valuemed3, valuemed4, valuemed5, loadingarea);
            }
        });

        review.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                try {
                    Desktop.getDesktop().open(new java.io.File("./Sortie.doc"));
                } catch (IOException e1) {
                    e1.printStackTrace();
                }
            }
        });

        clear.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                jtf1.setText("");
                jtf2.setText("");
                jtf3.setText("");
                jtf4.setText("");
                jtf5.setText("");
                loadingarea.setText("");
            }
        });

        exit.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                System.exit(0);
            }
        });
    }

    /**
     * Crée et affiche l'interface graphique.
     */
    private static void createAndShowGUI() {
        //Create and set up the window.
        reworkuniface frame = new reworkuniface("Rework UNIFACE");
        frame.setMinimumSize(new Dimension(780, 300));
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        //Set up the content pane.
        frame.addComponentsToPane(frame.getContentPane());
        //Display the window.
        frame.pack();
        frame.setLocationRelativeTo(null);
        frame.setVisible(true);
    }

    /**
     * Méthode principale pour exécuter le programme.
     */
    public static void main(String[] args) {
        try {
            UIManager.setLookAndFeel("javax.swing.plaf.metal.MetalLookAndFeel");
        } catch (UnsupportedLookAndFeelException ex) {
            ex.printStackTrace();
        } catch (IllegalAccessException ex) {
            ex.printStackTrace();
        } catch (InstantiationException ex) {
            ex.printStackTrace();
        } catch (ClassNotFoundException ex) {
            ex.printStackTrace();
        }
        UIManager.put("swing.boldMetal", Boolean.FALSE);
        javax.swing.SwingUtilities.invokeLater(new Runnable() {
            public void run() {
                createAndShowGUI();
            }
        });
    }

    public void itemStateChanged(ItemEvent e) {
    }

    public void valueChanged(ListSelectionEvent e) {
    }
}
