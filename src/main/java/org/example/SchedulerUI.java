package org.example;

import javax.swing.*;
import javax.swing.border.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.text.*;
import java.awt.*;
import java.awt.event.*;
import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;

public class SchedulerUI extends JFrame {

    // --- Data input publik ---------------------------------------------------
    public static String  inputFilePath  = "";
    public static String  outputFilePath = "output-jadwal.xlsx";
    public static int     kelasT7        = 8;
    public static int     kelasT8        = 8;
    public static int     kelasT9        = 7;
    public static int     jumlahKelas    = 23;
    public static int     jamSenin       = 10;
    public static int     jamSelasa      = 10;
    public static int     jamRabu        = 10;
    public static int     jamKamis       = 9;
    public static int     jamJumat       = 8;

    public static volatile boolean stopRequested = false;

    private static SchedulerUI   instance;
    private static SchedulerTask schedulerTask;
    private        Thread        schedulerThread;

    @FunctionalInterface
    public interface SchedulerTask { void run() throws Exception; }

    // --- Warna ---------------------------------------------------------------
    private static final Color C_BG        = new Color(245, 246, 248);
    private static final Color C_WHITE     = Color.WHITE;
    private static final Color C_BORDER    = new Color(220, 224, 230);
    private static final Color C_TEXT      = new Color(17, 24, 39);
    private static final Color C_SUBTEXT   = new Color(75, 85, 99);
    private static final Color C_MUTED     = new Color(156, 163, 175);
    private static final Color C_PRIMARY   = new Color(37, 99, 235);
    private static final Color C_INFO_BG   = new Color(239, 246, 255);
    private static final Color C_INFO_TEXT = new Color(29, 78, 216);
    private static final Color C_SUCCESS   = new Color(16, 185, 129);
    private static final Color C_SUCCESS_BG= new Color(236, 253, 245);
    private static final Color C_ERROR     = new Color(220, 38, 38);
    private static final Color C_WARN      = new Color(217, 119, 6);
    private static final Color C_DANGER_BG = new Color(254, 242, 242);
    private static final Color C_SPIN_BTN  = new Color(232, 236, 241);

    // --- Komponen ------------------------------------------------------------
    private JTextField   tfInputFile;
    private JLabel       lblInputDir;
    private JTextField   tfOutputFile;
    private JLabel       lblOutputInfo;
    private JLabel       lblFileChip;
    private JPanel       pnlFileChip;
    private JButton      btnInputFile;
    private JButton      btnOutputFile;
    // Spinner kelas (custom: minus/label/plus)
    private int[]        kelasVals = {8, 8, 7};
    private JLabel[]     kelasValLabel = new JLabel[3];
    // Spinner hari (custom)
    private int[]        hariVals = {10, 10, 10, 9, 8};
    private JLabel[]     hariValLabel = new JLabel[5];

    private JLabel       lblTotalKelas;
    private JButton      btnGenerate;
    private JButton      btnStop;
    private JTextPane    taLog;
    private JLabel       lblStatus;
    private JProgressBar progressBar;

    private java.util.List<JComponent> inputComponents = new java.util.ArrayList<>();

    // Log colors
    private static final Color LOG_OK   = new Color(52, 168, 83);
    private static final Color LOG_INFO = new Color(66, 133, 244);
    private static final Color LOG_WARN = new Color(251, 188, 4);
    private static final Color LOG_ERR  = new Color(234, 67, 53);
    private static final Color LOG_TEXT = new Color(200, 210, 220);
    private static final Color LOG_TIME = new Color(100, 115, 130);
    private static final Color LOG_BG   = new Color(13, 17, 23);

    // --- Konstruktor ---------------------------------------------------------
    private SchedulerUI() {
        super("Generator Jadwal Sekolah");
        instance = this;

        setDefaultCloseOperation(JFrame.DO_NOTHING_ON_CLOSE);
        addWindowListener(new WindowAdapter() {
            @Override public void windowClosing(WindowEvent e) {
                if (schedulerThread != null && schedulerThread.isAlive()) {
                    int r = JOptionPane.showConfirmDialog(SchedulerUI.this,
                            "Proses masih berjalan!\nStop proses dan keluar?",
                            "Konfirmasi Keluar", JOptionPane.YES_NO_OPTION,
                            JOptionPane.QUESTION_MESSAGE);
                    if (r != JOptionPane.YES_OPTION) return;
                    stopRequested = true;
                }
                System.exit(0);
            }
        });

        setSize(820, 800);
        setMinimumSize(new Dimension(700, 640));
        setLocationRelativeTo(null);
        getContentPane().setBackground(C_BG);

        buildUI();
        redirectOutput();
    }

    // --- Redirect output -----------------------------------------------------
    private void redirectOutput() {
        PrintStream ps = new PrintStream(new OutputStream() {
            private final StringBuilder buf = new StringBuilder();
            @Override public void write(int b) {
                char c = (char) b;
                buf.append(c);
                if (c == '\n') {
                    String line = buf.toString().stripTrailing();
                    buf.setLength(0);
                    SwingUtilities.invokeLater(() -> appendLog(line, null));
                }
            }
        }, true);
        System.setOut(ps);
        System.setErr(ps);
    }

    // --- Build UI ------------------------------------------------------------
    private void buildUI() {
        setLayout(new BorderLayout());
        add(buildHeader(), BorderLayout.NORTH);

        JPanel center = new JPanel();
        center.setLayout(new BoxLayout(center, BoxLayout.Y_AXIS));
        center.setBackground(C_BG);
        center.setBorder(BorderFactory.createEmptyBorder(14, 18, 14, 18));

        center.add(buildFileCard());
        center.add(box(10));
        center.add(buildKonfigCard());
        center.add(box(10));
        center.add(buildActionCard());
        center.add(box(10));
        center.add(buildLogCard());

        JScrollPane scroll = new JScrollPane(center);
        scroll.setBorder(null);
        scroll.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);
        scroll.getVerticalScrollBar().setUnitIncrement(16);
        scroll.setBackground(C_BG);
        add(scroll, BorderLayout.CENTER);
        add(buildStatusBar(), BorderLayout.SOUTH);
    }

    // -- Header ---------------------------------------------------------------
    private JPanel buildHeader() {
        JPanel p = new JPanel(new BorderLayout());
        p.setBackground(C_WHITE);
        p.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createMatteBorder(0, 0, 1, 0, C_BORDER),
                BorderFactory.createEmptyBorder(14, 20, 14, 20)));

        JPanel left = new JPanel(new FlowLayout(FlowLayout.LEFT, 12, 0));
        left.setOpaque(false);

        // Icon box
        JLabel iconBox = new JLabel("Jadwal");
        iconBox.setFont(new Font("Segoe UI", Font.BOLD, 12));
        iconBox.setOpaque(true);
        iconBox.setBackground(C_INFO_BG);
        iconBox.setForeground(C_INFO_TEXT);
        iconBox.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createLineBorder(new Color(191, 219, 254), 1, true),
                BorderFactory.createEmptyBorder(6, 8, 6, 8)));
        left.add(iconBox);

        JPanel titles = new JPanel();
        titles.setLayout(new BoxLayout(titles, BoxLayout.Y_AXIS));
        titles.setOpaque(false);
        JLabel title = new JLabel("Generator Jadwal Sekolah");
        title.setFont(new Font("Segoe UI", Font.BOLD, 15));
        title.setForeground(C_TEXT);
        JLabel sub = new JLabel("Hill Climbing Algorithm  \u00B7  Penjadwalan otomatis");
        sub.setFont(new Font("Segoe UI", Font.PLAIN, 11));
        sub.setForeground(C_MUTED);
        titles.add(title);
        titles.add(sub);
        left.add(titles);
        p.add(left, BorderLayout.WEST);

        return p;
    }

    // -- Card wrapper ---------------------------------------------------------
    private JPanel card(JPanel content) {
        JPanel card = new JPanel(new BorderLayout());
        card.setBackground(C_WHITE);
        card.setAlignmentX(Component.LEFT_ALIGNMENT);
        card.setMaximumSize(new Dimension(Integer.MAX_VALUE, Integer.MAX_VALUE));
        card.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createLineBorder(C_BORDER, 1, true),
                BorderFactory.createEmptyBorder(18, 20, 18, 20)));
        card.add(content);
        return card;
    }

    // -- Section label --------------------------------------------------------
    private JLabel sectionLabel(String text) {
        JLabel l = new JLabel(text.toUpperCase());
        l.setFont(new Font("Segoe UI", Font.BOLD, 10));
        l.setForeground(C_MUTED);
        l.setBorder(BorderFactory.createEmptyBorder(0, 0, 10, 0));
        l.setAlignmentX(Component.LEFT_ALIGNMENT);
        return l;
    }

    // -- File card ------------------------------------------------------------
    private JPanel buildFileCard() {
        JPanel p = new JPanel();
        p.setLayout(new BoxLayout(p, BoxLayout.Y_AXIS));
        p.setOpaque(false);
        p.add(sectionLabel("File Input & Output"));

        // ── Input row ──
        p.add(fieldLabel("File Excel input", true));
        p.add(box(4));

        // Direktori awal
        lblInputDir = dirLabel("Direktori: " + System.getProperty("user.dir"));
        p.add(lblInputDir);
        p.add(box(5));

        JPanel inputRow = hRow(8);
        tfInputFile = styledField("Belum ada file dipilih", true);
        tfInputFile.setEditable(false);
        tfInputFile.setForeground(C_MUTED);
        inputRow.add(tfInputFile);
        btnInputFile = smallBtn("Pilih file...");
        btnInputFile.addActionListener(e -> browseInput());
        inputRow.add(btnInputFile);
        p.add(inputRow);
        inputComponents.add(btnInputFile);

        // File chip (hidden initially)
        pnlFileChip = new JPanel(new FlowLayout(FlowLayout.LEFT, 0, 0));
        pnlFileChip.setOpaque(false);
        pnlFileChip.setAlignmentX(Component.LEFT_ALIGNMENT);
        lblFileChip = new JLabel("");
        lblFileChip.setFont(new Font("Segoe UI", Font.PLAIN, 11));
        lblFileChip.setForeground(new Color(6, 120, 80));
        lblFileChip.setBackground(C_SUCCESS_BG);
        lblFileChip.setOpaque(true);
        lblFileChip.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createLineBorder(new Color(167, 243, 208), 1, true),
                BorderFactory.createEmptyBorder(3, 9, 3, 9)));
        lblFileChip.setVisible(false);
        pnlFileChip.add(lblFileChip);
        p.add(box(5));
        p.add(pnlFileChip);

        p.add(box(14));

        // ── Output row ──
        p.add(fieldLabel("File output (.xlsx)", false));
        p.add(box(4));

        // Info lokasi output
        lblOutputInfo = dirLabel("Lokasi: " + new File("output-jadwal.xlsx").getAbsoluteFile().getParent()
                + File.separator + "output-jadwal.xlsx");
        p.add(lblOutputInfo);
        p.add(box(5));

        JPanel outputRow = hRow(8);
        tfOutputFile = styledField("output-jadwal.xlsx", false);
        tfOutputFile.getDocument().addDocumentListener(new javax.swing.event.DocumentListener() {
            public void insertUpdate(javax.swing.event.DocumentEvent e)  { refreshOutputInfo(); }
            public void removeUpdate(javax.swing.event.DocumentEvent e)  { refreshOutputInfo(); }
            public void changedUpdate(javax.swing.event.DocumentEvent e) { refreshOutputInfo(); }
        });
        outputRow.add(tfOutputFile);
        btnOutputFile = smallBtn("Lokasi simpan...");
        btnOutputFile.addActionListener(e -> browseOutput());
        outputRow.add(btnOutputFile);
        p.add(outputRow);
        inputComponents.add(tfOutputFile);
        inputComponents.add(btnOutputFile);

        return card(p);
    }

    /** Label untuk menampilkan path direktori. */
    private JLabel dirLabel(String text) {
        JLabel l = new JLabel(text);
        l.setFont(new Font("Segoe UI", Font.PLAIN, 13));
        l.setForeground(C_SUBTEXT);
        l.setAlignmentX(Component.LEFT_ALIGNMENT);
        return l;
    }

    /** Perbarui label info lokasi output sesuai isi text field. */
    private void refreshOutputInfo() {
        String raw = tfOutputFile.getText().trim();
        if (raw.isEmpty()) {
            lblOutputInfo.setText("Lokasi: (belum ditentukan)");
            return;
        }
        File f = new File(raw);
        if (!f.isAbsolute()) f = new File(System.getProperty("user.dir"), raw);
        String dir  = f.getAbsoluteFile().getParent();
        String name = f.getName();
        lblOutputInfo.setText("Lokasi: " + dir + File.separator + name);
    }

    // -- Konfigurasi card -----------------------------------------------------
    private JPanel buildKonfigCard() {
        JPanel p = new JPanel();
        p.setLayout(new BoxLayout(p, BoxLayout.Y_AXIS));
        p.setOpaque(false);
        p.add(sectionLabel("Konfigurasi Jadwal"));

        // Kelas per tingkat
        p.add(fieldLabel("Jumlah kelas per tingkat", false));
        p.add(box(8));

        String[] kelasNames  = {"Kelas 7", "Kelas 8", "Kelas 9"};
        Color[]  kelasColors = {new Color(37, 99, 235), new Color(5, 150, 105), new Color(180, 117, 23)};

        JPanel kelasGrid = new JPanel(new GridLayout(1, 3, 8, 0));
        kelasGrid.setOpaque(false);
        kelasGrid.setAlignmentX(Component.LEFT_ALIGNMENT);
        kelasGrid.setMaximumSize(new Dimension(Integer.MAX_VALUE, 75));

        for (int i = 0; i < 3; i++) {
            final int idx = i;
            kelasValLabel[i] = spinLabel(String.valueOf(kelasVals[i]));
            JButton minus = spinBtn("-");
            JButton plus  = spinBtn("+");
            minus.addActionListener(e -> { kelasVals[idx] = Math.max(0, kelasVals[idx]-1); kelasValLabel[idx].setText(String.valueOf(kelasVals[idx])); updateTotalKelas(); });
            plus .addActionListener(e -> { kelasVals[idx] = Math.min(50, kelasVals[idx]+1); kelasValLabel[idx].setText(String.valueOf(kelasVals[idx])); updateTotalKelas(); });
            inputComponents.add(minus); inputComponents.add(plus);
            kelasGrid.add(buildSpinnerCard(kelasNames[i], kelasColors[i], minus, kelasValLabel[i], plus));
        }
        p.add(kelasGrid);
        p.add(box(8));

        // Total kelas info row
        lblTotalKelas = new JLabel("  Total 23 kelas  (7A-7H, 8A-8H, 9A-9G)");
        lblTotalKelas.setFont(new Font("Segoe UI", Font.PLAIN, 12));
        lblTotalKelas.setForeground(C_INFO_TEXT);
        lblTotalKelas.setBackground(C_INFO_BG);
        lblTotalKelas.setOpaque(true);
        lblTotalKelas.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createLineBorder(new Color(191, 219, 254), 1, true),
                BorderFactory.createEmptyBorder(6, 12, 6, 12)));
        lblTotalKelas.setAlignmentX(Component.LEFT_ALIGNMENT);
        lblTotalKelas.setMaximumSize(new Dimension(Integer.MAX_VALUE, 36));
        p.add(lblTotalKelas);
        p.add(box(16));

        // Jam per hari
        p.add(fieldLabel("Jam pelajaran per hari", false));
        p.add(box(8));

        String[] hariNames = {"Senin", "Selasa", "Rabu", "Kamis", "Jumat"};

        JPanel hariGrid = new JPanel(new GridLayout(1, 5, 8, 0));
        hariGrid.setOpaque(false);
        hariGrid.setAlignmentX(Component.LEFT_ALIGNMENT);
        hariGrid.setMaximumSize(new Dimension(Integer.MAX_VALUE, 75));

        for (int i = 0; i < 5; i++) {
            final int idx = i;
            hariValLabel[i] = spinLabel(String.valueOf(hariVals[i]));
            JButton minus = spinBtn("-");
            JButton plus  = spinBtn("+");
            minus.addActionListener(e -> { hariVals[idx] = Math.max(1, hariVals[idx]-1); hariValLabel[idx].setText(String.valueOf(hariVals[idx])); });
            plus .addActionListener(e -> { hariVals[idx] = Math.min(20, hariVals[idx]+1); hariValLabel[idx].setText(String.valueOf(hariVals[idx])); });
            inputComponents.add(minus); inputComponents.add(plus);
            hariGrid.add(buildSpinnerCard(hariNames[i], C_SUBTEXT, minus, hariValLabel[i], plus));
        }
        p.add(hariGrid);

        return card(p);
    }

    // -- Action card ----------------------------------------------------------
    private JPanel buildActionCard() {
        JPanel p = new JPanel(new BorderLayout(10, 0));
        p.setOpaque(false);
        p.setAlignmentX(Component.LEFT_ALIGNMENT);

        btnGenerate = new JButton("Generate Jadwal");
        btnGenerate.setFont(new Font("Segoe UI", Font.BOLD, 14));
        btnGenerate.setBackground(C_PRIMARY);
        btnGenerate.setForeground(Color.WHITE);
        btnGenerate.setOpaque(true);
        btnGenerate.setFocusPainted(false);
        btnGenerate.setBorderPainted(false);
        btnGenerate.setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
        btnGenerate.setPreferredSize(new Dimension(0, 42));
        btnGenerate.addActionListener(e -> onGenerate());
        btnGenerate.addMouseListener(hoverEffect(btnGenerate, C_PRIMARY, new Color(29, 78, 216)));

        btnStop = new JButton("Stop & Simpan");
        btnStop.setFont(new Font("Segoe UI", Font.BOLD, 13));
        btnStop.setBackground(C_DANGER_BG);
        btnStop.setForeground(C_ERROR);
        btnStop.setOpaque(true);
        btnStop.setFocusPainted(false);
        btnStop.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createLineBorder(new Color(252, 165, 165), 1, true),
                BorderFactory.createEmptyBorder(0, 16, 0, 16)));
        btnStop.setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
        btnStop.setPreferredSize(new Dimension(160, 42));
        btnStop.setVisible(false);
        btnStop.addActionListener(e -> onStop());

        progressBar = new JProgressBar();
        progressBar.setIndeterminate(true);
        progressBar.setVisible(false);
        progressBar.setPreferredSize(new Dimension(0, 3));
        progressBar.setBackground(C_INFO_BG);
        progressBar.setForeground(C_PRIMARY);
        progressBar.setBorderPainted(false);

        JPanel btnRow = new JPanel(new BorderLayout(10, 0));
        btnRow.setOpaque(false);
        btnRow.add(btnGenerate, BorderLayout.CENTER);
        btnRow.add(btnStop, BorderLayout.EAST);

        JPanel wrapper = new JPanel(new BorderLayout());
        wrapper.setBackground(C_WHITE);
        wrapper.setAlignmentX(Component.LEFT_ALIGNMENT);
        wrapper.setMaximumSize(new Dimension(Integer.MAX_VALUE, Integer.MAX_VALUE));
        wrapper.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createLineBorder(C_BORDER, 1, true),
                BorderFactory.createEmptyBorder(14, 20, 14, 20)));
        wrapper.add(btnRow, BorderLayout.CENTER);
        wrapper.add(progressBar, BorderLayout.SOUTH);
        return wrapper;
    }

    // -- Log card -------------------------------------------------------------
    private JPanel buildLogCard() {
        JPanel outer = new JPanel();
        outer.setLayout(new BoxLayout(outer, BoxLayout.Y_AXIS));
        outer.setOpaque(false);

        // Header row
        JPanel logHeader = new JPanel(new BorderLayout());
        logHeader.setOpaque(false);
        logHeader.setAlignmentX(Component.LEFT_ALIGNMENT);
        logHeader.setMaximumSize(new Dimension(Integer.MAX_VALUE, 30));
        JLabel logTitle = new JLabel("Log Proses");
        logTitle.setFont(new Font("Segoe UI", Font.BOLD, 12));
        logTitle.setForeground(C_SUBTEXT);
        logHeader.add(logTitle, BorderLayout.WEST);

        JButton btnClear = new JButton("Bersihkan");
        btnClear.setFont(new Font("Segoe UI", Font.PLAIN, 11));
        btnClear.setForeground(C_MUTED);
        btnClear.setBackground(C_BG);
        btnClear.setFocusPainted(false);
        btnClear.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createLineBorder(C_BORDER, 1, true),
                BorderFactory.createEmptyBorder(2, 9, 2, 9)));
        btnClear.setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
        btnClear.addActionListener(e -> {
            taLog.setText("");
            appendLog("  Log dibersihkan.", LOG_TEXT);
        });
        logHeader.add(btnClear, BorderLayout.EAST);

        outer.add(logHeader);
        outer.add(box(8));

        // Log pane
        taLog = new JTextPane();
        taLog.setEditable(false);
        taLog.setBackground(LOG_BG);
        taLog.setForeground(LOG_TEXT);
        taLog.setFont(new Font("Consolas", Font.PLAIN, 12));
        taLog.setBorder(BorderFactory.createEmptyBorder(12, 14, 12, 14));
        appendLog("  Selamat datang! Pilih file input dan atur konfigurasi, lalu klik Generate.", LOG_TEXT);

        JScrollPane logScroll = new JScrollPane(taLog);
        logScroll.setBorder(BorderFactory.createLineBorder(C_BORDER, 1, true));
        logScroll.setPreferredSize(new Dimension(0, 200));
        logScroll.setAlignmentX(Component.LEFT_ALIGNMENT);
        outer.add(logScroll);

        JPanel wrapper = new JPanel(new BorderLayout());
        wrapper.setBackground(C_WHITE);
        wrapper.setAlignmentX(Component.LEFT_ALIGNMENT);
        wrapper.setMaximumSize(new Dimension(Integer.MAX_VALUE, Integer.MAX_VALUE));
        wrapper.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createLineBorder(C_BORDER, 1, true),
                BorderFactory.createEmptyBorder(14, 20, 14, 20)));
        wrapper.add(outer, BorderLayout.CENTER);
        return wrapper;
    }

    // -- Status bar -----------------------------------------------------------
    private JPanel buildStatusBar() {
        JPanel bar = new JPanel(new BorderLayout());
        bar.setBackground(C_WHITE);
        bar.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createMatteBorder(1, 0, 0, 0, C_BORDER),
                BorderFactory.createEmptyBorder(7, 20, 7, 20)));

        lblStatus = new JLabel("Siap");
        lblStatus.setFont(new Font("Segoe UI", Font.PLAIN, 12));
        lblStatus.setForeground(C_SUCCESS);
        bar.add(lblStatus, BorderLayout.WEST);

        JLabel hint = new JLabel("Klik Stop untuk simpan hasil terbaik saat ini");
        hint.setFont(new Font("Segoe UI", Font.PLAIN, 11));
        hint.setForeground(C_MUTED);
        bar.add(hint, BorderLayout.EAST);
        return bar;
    }

    // --- Helper builders -----------------------------------------------------

    private JTextField styledField(String placeholder, boolean readOnly) {
        JTextField tf = new JTextField(placeholder);
        tf.setFont(new Font("Segoe UI", Font.PLAIN, 13));
        tf.setForeground(readOnly ? C_MUTED : C_TEXT);
        tf.setBackground(readOnly ? C_BG : C_WHITE);
        tf.setEditable(!readOnly);
        tf.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createLineBorder(C_BORDER, 1, true),
                BorderFactory.createEmptyBorder(8, 11, 8, 11)));
        if (!readOnly) {
            tf.addFocusListener(new FocusAdapter() {
                @Override public void focusGained(FocusEvent e) {
                    tf.setBorder(BorderFactory.createCompoundBorder(
                            BorderFactory.createLineBorder(C_PRIMARY, 1, true),
                            BorderFactory.createEmptyBorder(8, 11, 8, 11)));
                }
                @Override public void focusLost(FocusEvent e) {
                    tf.setBorder(BorderFactory.createCompoundBorder(
                            BorderFactory.createLineBorder(C_BORDER, 1, true),
                            BorderFactory.createEmptyBorder(8, 11, 8, 11)));
                }
            });
        }
        return tf;
    }

    private JButton smallBtn(String text) {
        JButton b = new JButton(text);
        b.setFont(new Font("Segoe UI", Font.PLAIN, 12));
        b.setBackground(C_WHITE);
        b.setForeground(C_SUBTEXT);
        b.setFocusPainted(false);
        b.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createLineBorder(C_BORDER, 1, true),
                BorderFactory.createEmptyBorder(8, 13, 8, 13)));
        b.setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
        b.setPreferredSize(new Dimension(140, 38));
        b.addMouseListener(new MouseAdapter() {
            @Override public void mouseEntered(MouseEvent e) { b.setBackground(C_BG); }
            @Override public void mouseExited(MouseEvent e)  { b.setBackground(C_WHITE); }
        });
        return b;
    }

    private JButton spinBtn(String text) {
        JButton b = new JButton(text);
        b.setFont(new Font("Segoe UI", Font.BOLD, 15));
        b.setBackground(C_SPIN_BTN);
        b.setForeground(C_TEXT);
        b.setFocusPainted(false);
        b.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createLineBorder(C_BORDER, 1, true),
                BorderFactory.createEmptyBorder(0, 5, 0, 5)));
        b.setOpaque(true);
        b.setPreferredSize(new Dimension(28, 26));
        b.setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
        b.addMouseListener(new MouseAdapter() {
            @Override public void mouseEntered(MouseEvent e) { b.setBackground(new Color(209, 213, 219)); }
            @Override public void mouseExited(MouseEvent e)  { b.setBackground(C_SPIN_BTN); }
            @Override public void mousePressed(MouseEvent e) { b.setBackground(new Color(165, 170, 180)); }
            @Override public void mouseReleased(MouseEvent e) { b.setBackground(C_SPIN_BTN); }
        });
        return b;
    }

    private JLabel spinLabel(String val) {
        JLabel l = new JLabel(val, SwingConstants.CENTER);
        l.setFont(new Font("Segoe UI", Font.BOLD, 14));
        l.setForeground(C_TEXT);
        l.setPreferredSize(new Dimension(28, 26));
        return l;
    }

    private JPanel buildSpinnerCard(String name, Color dotColor, JButton minus, JLabel valLbl, JButton plus) {
        JPanel card = new JPanel(new BorderLayout(0, 6));
        card.setBackground(C_BG);
        card.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createLineBorder(C_BORDER, 1, true),
                BorderFactory.createEmptyBorder(10, 8, 10, 8)));

        // Name row with colored dot
        JPanel nameRow = new JPanel(new FlowLayout(FlowLayout.CENTER, 4, 0));
        nameRow.setOpaque(false);
        JLabel dot = new JLabel("*");
        dot.setFont(new Font("Segoe UI", Font.BOLD, 10));
        dot.setForeground(dotColor);
        JLabel nameLbl = new JLabel(name);
        nameLbl.setFont(new Font("Segoe UI", Font.PLAIN, 11));
        nameLbl.setForeground(C_SUBTEXT);
        nameRow.add(dot);
        nameRow.add(nameLbl);

        // Spinner control row
        JPanel ctlRow = new JPanel(new FlowLayout(FlowLayout.CENTER, 4, 0));
        ctlRow.setOpaque(false);
        ctlRow.add(minus);
        ctlRow.add(valLbl);
        ctlRow.add(plus);

        card.add(nameRow, BorderLayout.NORTH);
        card.add(ctlRow, BorderLayout.CENTER);
        return card;
    }

    private JLabel fieldLabel(String text, boolean required) {
        JLabel l = new JLabel(text + (required ? " *" : ""));
        l.setFont(new Font("Segoe UI", Font.BOLD, 12));
        l.setForeground(required ? C_TEXT : C_SUBTEXT);
        l.setAlignmentX(Component.LEFT_ALIGNMENT);
        return l;
    }

    private JPanel hRow(int gap) {
        JPanel row = new JPanel(new BorderLayout(gap, 0));
        row.setOpaque(false);
        row.setAlignmentX(Component.LEFT_ALIGNMENT);
        row.setMaximumSize(new Dimension(Integer.MAX_VALUE, 42));
        return row;
    }

    private Component box(int h) { return Box.createVerticalStrut(h); }

    private MouseAdapter hoverEffect(JButton btn, Color normal, Color hover) {
        return new MouseAdapter() {
            @Override public void mouseEntered(MouseEvent e) { if (btn.isEnabled()) btn.setBackground(hover); }
            @Override public void mouseExited(MouseEvent e)  { if (btn.isEnabled()) btn.setBackground(normal); }
        };
    }

    private void updateTotalKelas() {
        int t7 = kelasVals[0], t8 = kelasVals[1], t9 = kelasVals[2];
        int total = t7 + t8 + t9;
        String p7 = t7 == 0 ? "-" : "7A" + (t7 > 1 ? "-7" + (char)('A'+t7-1) : "");
        String p8 = t8 == 0 ? "-" : "8A" + (t8 > 1 ? "-8" + (char)('A'+t8-1) : "");
        String p9 = t9 == 0 ? "-" : "9A" + (t9 > 1 ? "-9" + (char)('A'+t9-1) : "");
        lblTotalKelas.setText("  Total " + total + " kelas  (" + p7 + ", " + p8 + ", " + p9 + ")");
    }

    // --- Colored log ---------------------------------------------------------
    private void appendLog(String text, Color color) {
        StyledDocument doc = taLog.getStyledDocument();
        Style style = taLog.addStyle("s", null);

        // Timestamp
        StyleConstants.setForeground(style, LOG_TIME);
        StyleConstants.setFontFamily(style, "Consolas");
        StyleConstants.setFontSize(style, 12);
        String ts = new SimpleDateFormat("HH:mm:ss").format(new Date());
        try {
            doc.insertString(doc.getLength(), ts + "  ", style);
        } catch (BadLocationException ignored) {}

        // Message
        Color msgColor = color != null ? color : detectLogColor(text);
        StyleConstants.setForeground(style, msgColor);
        try {
            doc.insertString(doc.getLength(), text + "\n", style);
        } catch (BadLocationException ignored) {}

        taLog.setCaretPosition(doc.getLength());
    }

    private Color detectLogColor(String line) {
        if (line == null) return LOG_TEXT;
        String l = line.toLowerCase();
        if (l.contains("selesai") || l.contains("berhasil")) return LOG_OK;
        if (l.contains("gagal")   || l.contains("error"))    return LOG_ERR;
        if (l.contains("konflik") || l.contains("peringatan") || l.contains("stop")) return LOG_WARN;
        if (l.contains("iterasi") || l.contains("->"))        return LOG_INFO;
        return LOG_TEXT;
    }

    // --- Aksi ----------------------------------------------------------------
    private void browseInput() {
        JFileChooser fc = new JFileChooser();
        fc.setDialogTitle("Pilih File Excel Input");
        fc.setFileFilter(new FileNameExtensionFilter("Excel Workbook (*.xlsx)", "xlsx"));
        // Mulai dari direktori kerja atau direktori file input sebelumnya
        String current = tfInputFile.getText().trim();
        if (!current.isEmpty() && !current.startsWith("Belum")) {
            File prev = new File(current);
            if (prev.exists()) fc.setCurrentDirectory(prev.getParentFile());
        } else {
            fc.setCurrentDirectory(new File(System.getProperty("user.dir")));
        }
        if (fc.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            File f = fc.getSelectedFile();
            tfInputFile.setText(f.getAbsolutePath());
            tfInputFile.setForeground(C_TEXT);
            long kb = f.length() / 1024;
            // Update label direktori input
            lblInputDir.setText("Direktori: " + f.getParent());
            lblInputDir.setForeground(C_INFO_TEXT);
            // Chip
            lblFileChip.setText("[OK]  " + f.getName() + "  -  " + kb + " KB");
            lblFileChip.setVisible(true);
        }
    }

    private void browseOutput() {
        JFileChooser fc = new JFileChooser();
        fc.setDialogTitle("Pilih Lokasi Simpan Output");
        fc.setFileFilter(new FileNameExtensionFilter("Excel Workbook (*.xlsx)", "xlsx"));
        String raw = tfOutputFile.getText().trim();
        File suggested = raw.isEmpty() ? new File("output-jadwal.xlsx") : new File(raw);
        if (!suggested.isAbsolute()) suggested = new File(System.getProperty("user.dir"), suggested.getName());
        fc.setCurrentDirectory(suggested.getParentFile());
        fc.setSelectedFile(suggested);
        if (fc.showSaveDialog(this) == JFileChooser.APPROVE_OPTION) {
            String path = fc.getSelectedFile().getAbsolutePath();
            if (!path.endsWith(".xlsx")) path += ".xlsx";
            tfOutputFile.setText(path);
            // refreshOutputInfo() dipanggil otomatis via DocumentListener
        }
    }

    private void onGenerate() {
        String inPath = tfInputFile.getText().trim();
        if (inPath.isEmpty() || inPath.startsWith("Belum")) {
            showError("File input belum dipilih!", "Pilih file Excel input terlebih dahulu.");
            return;
        }
        if (!new File(inPath).exists()) {
            showError("File tidak ditemukan", "File berikut tidak ada:\n" + inPath);
            return;
        }
        int totalK = kelasVals[0] + kelasVals[1] + kelasVals[2];
        if (totalK == 0) {
            showError("Kelas kosong", "Jumlah total kelas tidak boleh 0.");
            return;
        }

        inputFilePath  = inPath;
        outputFilePath = tfOutputFile.getText().trim().isEmpty() ? "output-jadwal.xlsx" : tfOutputFile.getText().trim();
        kelasT7     = kelasVals[0];
        kelasT8     = kelasVals[1];
        kelasT9     = kelasVals[2];
        jumlahKelas = kelasT7 + kelasT8 + kelasT9;
        jamSenin    = hariVals[0];
        jamSelasa   = hariVals[1];
        jamRabu     = hariVals[2];
        jamKamis    = hariVals[3];
        jamJumat    = hariVals[4];

        stopRequested = false;
        setProcessing(true);

        String ts = new SimpleDateFormat("HH:mm:ss").format(new Date());
        appendLog("=".repeat(52), C_MUTED);
        appendLog("  Generate dimulai  [" + ts + "]", LOG_INFO);
        appendLog("  Input   : " + new File(inputFilePath).getName(), LOG_TEXT);
        appendLog("  Output  : " + outputFilePath, LOG_TEXT);
        appendLog("  Kelas   : " + jumlahKelas + " total  (7=" + kelasT7 + ", 8=" + kelasT8 + ", 9=" + kelasT9 + ")", LOG_TEXT);
        appendLog(String.format("  Jam     : Sn=%d Se=%d Ra=%d Ka=%d Jm=%d", jamSenin, jamSelasa, jamRabu, jamKamis, jamJumat), LOG_TEXT);
        appendLog("-".repeat(52), C_MUTED);

        schedulerThread = new Thread(() -> {
            try {
                schedulerTask.run();
            } catch (Exception ex) {
                ex.printStackTrace();
                String msg = ex.getMessage();
                String exName = ex.getClass().getSimpleName();
                boolean excelError = msg != null && (
                        exName.contains("OfficeXml") ||
                        exName.contains("POI") ||
                        msg.toLowerCase().contains("excel") ||
                        msg.toLowerCase().contains("workbook") ||
                        msg.toLowerCase().contains(".xlsx") ||
                        ex instanceof java.io.FileNotFoundException ||
                        ex instanceof java.io.IOException);
                if (excelError) {
                    markDone(false, "EXCEL_ERROR:" + (msg != null ? msg : ""));
                } else {
                    markDone(false, "Exception: " + (msg != null ? msg : "Unknown error"));
                }
            }
        }, "SchedulerThread");
        schedulerThread.setDaemon(true);
        schedulerThread.start();
    }

    private void onStop() {
        btnStop.setEnabled(false);
        btnStop.setText("Menghentikan...");
        stopRequested = true;
        appendLog("  STOP diminta -- menunggu iterasi selesai lalu export...", LOG_WARN);
        setStatus("Menghentikan proses...", C_WARN);
    }

    private void setProcessing(boolean processing) {
        btnGenerate.setEnabled(!processing);
        btnGenerate.setText(processing ? "Sedang memproses..." : "Generate Jadwal");
        btnGenerate.setBackground(processing ? new Color(100, 116, 139) : C_PRIMARY);
        btnStop.setVisible(processing);
        btnStop.setEnabled(processing);
        btnStop.setText("Stop & Simpan");
        progressBar.setVisible(processing);

        for (JComponent comp : inputComponents) comp.setEnabled(!processing);
        if (processing) setStatus("Memproses jadwal...", C_WARN);
    }

    // --- Public API ----------------------------------------------------------
    public static void markDone(boolean success, String extraMessage) {
        SwingUtilities.invokeLater(() -> {
            if (instance == null) return;
            if (extraMessage != null && !extraMessage.isBlank())
                instance.appendLog(extraMessage, null);

            instance.setProcessing(false);
            stopRequested = false;

            String ts = new SimpleDateFormat("HH:mm:ss").format(new Date());

            if (success) {
                instance.appendLog("-".repeat(52), C_MUTED);
                instance.appendLog("  Selesai  [" + ts + "]", LOG_OK);
                instance.appendLog("  File : " + outputFilePath, LOG_TEXT);
                instance.appendLog("=".repeat(52), C_MUTED);
                instance.setStatus("Selesai  -  " + new File(outputFilePath).getName(), C_SUCCESS);

                Object[] options = {"Oke", "Buka Folder Output"};
                int choice = JOptionPane.showOptionDialog(instance,
                        "<html><div style='padding:8px'>"
                                + "<b style='font-size:14px;color:#059669'>Jadwal berhasil digenerate!</b>"
                                + "<br><br>File disimpan di:<br>"
                                + "<code style='font-size:11px'>" + outputFilePath + "</code>"
                                + "</div></html>",
                        "Berhasil", JOptionPane.YES_NO_OPTION, JOptionPane.INFORMATION_MESSAGE,
                        null, options, options[0]);

                if (choice == 1) {
                    try { Desktop.getDesktop().open(new File(outputFilePath).getAbsoluteFile().getParentFile()); }
                    catch (Exception ex) { JOptionPane.showMessageDialog(instance, "Tidak bisa membuka folder: " + ex.getMessage()); }
                }
            } else {
                instance.appendLog("-".repeat(52), C_MUTED);
                instance.appendLog("  Gagal / Dihentikan  [" + ts + "]", LOG_ERR);
                instance.appendLog("=".repeat(52), C_MUTED);
                instance.setStatus("Dihentikan  -  klik Generate untuk coba lagi", C_ERROR);

                boolean isExcelError = extraMessage != null && extraMessage.startsWith("EXCEL_ERROR:");
                if (isExcelError) {
                    JOptionPane.showMessageDialog(instance,
                            "<html><div style='padding:8px'>"
                                    + "<b style='font-size:14px;color:#dc2626'>File Excel Tidak Valid</b>"
                                    + "<br><br>"
                                    + "<div style='background:#FEF2F2;padding:10px;border-radius:6px;color:#991B1B;font-size:13px'>"
                                    + "Format file Excel atau file Excel salah."
                                    + "</div>"
                                    + "<br>Pastikan file yang dipilih adalah:<br>"
                                    + "\u2022 File .xlsx (bukan .xls)<br>"
                                    + "\u2022 Format sesuai template yang ditentukan"
                                    + "</div></html>",
                            "Gagal", JOptionPane.ERROR_MESSAGE);
                } else {
                    JOptionPane.showMessageDialog(instance,
                            "<html><div style='padding:8px'>"
                                    + "<b style='font-size:14px;color:#dc2626'>Generate gagal atau dihentikan</b>"
                                    + "<br><br>Cek log untuk detail."
                                    + "</div></html>",
                            "Gagal", JOptionPane.ERROR_MESSAGE);
                }
            }
        });
    }

    // --- Helpers -------------------------------------------------------------
    private void setStatus(String msg, Color color) {
        lblStatus.setText(msg);
        lblStatus.setForeground(color);
    }

    private void showError(String title, String message) {
        JOptionPane.showMessageDialog(this, message, title, JOptionPane.WARNING_MESSAGE);
    }

    // --- Entry point ---------------------------------------------------------
    public static void launch(SchedulerTask task) {
        schedulerTask = task;
        try { UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName()); }
        catch (Exception ignored) {}
        SwingUtilities.invokeLater(() -> new SchedulerUI().setVisible(true));
    }
}