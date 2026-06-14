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
import java.util.List;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableModel;
import java.awt.print.PrinterJob;
import java.awt.print.Printable;
import java.awt.print.PageFormat;

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

    // Data dari Main setelah generate selesai
    public static String[][]   scheduleJadwal;
    public static int[][]      scheduleHariRange;
    public static List<String[]> scheduleKebutuhan;

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
    private JPanel       pnlResult;

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

        pnlResult = buildResultCard();
        pnlResult.setVisible(false);
        center.add(pnlResult);
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
            plus .addActionListener(e -> { hariVals[idx] = Math.min(10, hariVals[idx]+1); hariValLabel[idx].setText(String.valueOf(hariVals[idx])); });
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

    // -- Result card (hidden initially, shown after generate) -----------------
    private JPanel buildResultCard() {
        JPanel outer = new JPanel();
        outer.setLayout(new BoxLayout(outer, BoxLayout.Y_AXIS));
        outer.setOpaque(false);
        outer.add(sectionLabel("Hasil Export"));

        // Tab-like container
        JTabbedPane tabs = new JTabbedPane();
        tabs.setFont(new Font("Segoe UI", Font.PLAIN, 12));
        tabs.setBackground(C_WHITE);

        // ── Per Guru tab ──
        JPanel guruPanel = new JPanel();
        guruPanel.setLayout(new BoxLayout(guruPanel, BoxLayout.Y_AXIS));
        guruPanel.setBackground(C_WHITE);
        JScrollPane guruScroll = new JScrollPane(guruPanel);
        guruScroll.setBorder(null);
        guruScroll.getVerticalScrollBar().setUnitIncrement(12);
        tabs.addTab("Per Guru", guruScroll);

        // ── Per Kelas tab ──
        JPanel kelasPanel = new JPanel();
        kelasPanel.setLayout(new BoxLayout(kelasPanel, BoxLayout.Y_AXIS));
        kelasPanel.setBackground(C_WHITE);
        JScrollPane kelasScroll = new JScrollPane(kelasPanel);
        kelasScroll.setBorder(null);
        kelasScroll.getVerticalScrollBar().setUnitIncrement(12);
        tabs.addTab("Per Kelas", kelasScroll);

        // Store references for later population
        tabs.setName("resultTabs");
        guruPanel.setName("guruPanel");
        kelasPanel.setName("kelasPanel");

        outer.add(tabs);

        JPanel wrapper = new JPanel(new BorderLayout());
        wrapper.setBackground(C_WHITE);
        wrapper.setAlignmentX(Component.LEFT_ALIGNMENT);
        wrapper.setMaximumSize(new Dimension(Integer.MAX_VALUE, 320));
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
            long kb = f.length() / 1024;
            int confirm = JOptionPane.showConfirmDialog(this,
                    "<html><div style='padding:6px'>"
                        + "<b>Konfirmasi file input</b><br><br>"
                        + "<b>Nama:</b> " + f.getName() + "<br>"
                        + "<b>Ukuran:</b> " + kb + " KB<br>"
                        + "<b>Lokasi:</b> " + f.getParent() + "<br><br>"
                        + "Apakah file ini sudah benar?"
                        + "</div></html>",
                    "Konfirmasi", JOptionPane.YES_NO_OPTION, JOptionPane.QUESTION_MESSAGE);
            if (confirm != JOptionPane.YES_OPTION) return;
            tfInputFile.setText(f.getAbsolutePath());
            tfInputFile.setForeground(C_TEXT);
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

        // Auto-rename jika file sudah ada
        File of = new File(outputFilePath);
        if (of.exists()) {
            String name = of.getName();
            int dot = name.lastIndexOf('.');
            String base = (dot > 0) ? name.substring(0, dot) : name;
            String ext  = (dot > 0) ? name.substring(dot) : "";
            String parent = of.getParent();
            if (parent == null) parent = ".";
            int n = 1;
            File nf;
            do {
                nf = new File(parent, base + " (" + n + ")" + ext);
                n++;
            } while (nf.exists());
            outputFilePath = nf.getAbsolutePath();
            tfOutputFile.setText(outputFilePath);
            appendLog("  File output sudah ada, rename: " + nf.getName(), LOG_WARN);
        }
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
        if (pnlResult != null) pnlResult.setVisible(false);

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

    // --- Data passing from Main ----------------------------------------------
    public static void setData(String[][] jadwal, int[][] hariRange, List<String[]> kebutuhan) {
        scheduleJadwal = jadwal;
        scheduleHariRange = hariRange;
        scheduleKebutuhan = kebutuhan;
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
                instance.showResultPanel();
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

    // --- Result panel ---------------------------------------------------------
    private void showResultPanel() {
        if (scheduleJadwal == null || scheduleKebutuhan == null) return;

        // Find the tabs and panels via name
        JTabbedPane tabs = null;
        for (java.awt.Component c : pnlResult.getComponents()) {
            if (c instanceof javax.swing.JPanel) {
                for (java.awt.Component c2 : ((javax.swing.JPanel)c).getComponents()) {
                    if (c2 instanceof JTabbedPane) { tabs = (JTabbedPane) c2; break; }
                }
            }
        }
        if (tabs == null) return;

        // ── Populate Per Guru tab ──
        JScrollPane guruScroll = (JScrollPane) tabs.getComponentAt(0);
        JPanel guruPanel = (JPanel) guruScroll.getViewport().getView();
        guruPanel.removeAll();

        // Collect unique teachers
        java.util.Map<String, String> guruMap = new java.util.LinkedHashMap<>();
        for (String[] data : scheduleKebutuhan) {
            String num = data[0].replaceAll("[^0-9]", "");
            if (!guruMap.containsKey(num))
                guruMap.put(num, data[1]);
        }
        java.util.List<String> guruList = new java.util.ArrayList<>(guruMap.keySet());
        guruList.sort(java.util.Comparator.comparingInt(Integer::parseInt));

        for (String guruNum : guruList) {
            String namaGuru = guruMap.get(guruNum);
            JPanel row = new JPanel(new BorderLayout(8, 0));
            row.setOpaque(false);
            row.setMaximumSize(new Dimension(Integer.MAX_VALUE, 36));
            row.setBorder(BorderFactory.createEmptyBorder(2, 6, 2, 6));

            JLabel lbl = new JLabel("  Guru " + guruNum + " - " + namaGuru);
            lbl.setFont(new Font("Segoe UI", Font.PLAIN, 12));
            lbl.setForeground(C_TEXT);

            JPanel btnPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT, 4, 0));
            btnPanel.setOpaque(false);

            JButton btnExcel = new JButton("Excel");
            btnExcel.setFont(new Font("Segoe UI", Font.PLAIN, 11));
            btnExcel.setBackground(C_WHITE);
            btnExcel.setForeground(C_PRIMARY);
            btnExcel.setFocusPainted(false);
            btnExcel.setBorder(BorderFactory.createCompoundBorder(
                    BorderFactory.createLineBorder(new Color(191, 219, 254), 1, true),
                    BorderFactory.createEmptyBorder(4, 10, 4, 10)));
            btnExcel.setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
            String gNum = guruNum;
            btnExcel.addActionListener(e -> onExportGuruExcel(gNum));

            JButton btnCetak = new JButton("Cetak");
            btnCetak.setFont(new Font("Segoe UI", Font.PLAIN, 11));
            btnCetak.setBackground(C_WHITE);
            btnCetak.setForeground(C_SUCCESS);
            btnCetak.setFocusPainted(false);
            btnCetak.setBorder(BorderFactory.createCompoundBorder(
                    BorderFactory.createLineBorder(new Color(167, 243, 208), 1, true),
                    BorderFactory.createEmptyBorder(4, 10, 4, 10)));
            btnCetak.setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
            btnCetak.addActionListener(e -> onPrintGuru(gNum, namaGuru));

            btnPanel.add(btnExcel);
            btnPanel.add(btnCetak);

            row.add(lbl, BorderLayout.CENTER);
            row.add(btnPanel, BorderLayout.EAST);
            guruPanel.add(row);
            guruPanel.add(Box.createVerticalStrut(2));
        }

        // ── Populate Per Kelas tab ──
        JScrollPane kelasScroll = (JScrollPane) tabs.getComponentAt(1);
        JPanel kelasPanel = (JPanel) kelasScroll.getViewport().getView();
        kelasPanel.removeAll();

        String[] allKelas = Main.buatDaftarKelas();
        for (int kIdx = 0; kIdx < allKelas.length; kIdx++) {
            String namaKelas = allKelas[kIdx];
            JPanel row = new JPanel(new BorderLayout(8, 0));
            row.setOpaque(false);
            row.setMaximumSize(new Dimension(Integer.MAX_VALUE, 36));
            row.setBorder(BorderFactory.createEmptyBorder(2, 6, 2, 6));

            JLabel lbl = new JLabel("  Kelas " + namaKelas);
            lbl.setFont(new Font("Segoe UI", Font.PLAIN, 12));
            lbl.setForeground(C_TEXT);

            JPanel btnPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT, 4, 0));
            btnPanel.setOpaque(false);

            JButton btnExcel = new JButton("Excel");
            btnExcel.setFont(new Font("Segoe UI", Font.PLAIN, 11));
            btnExcel.setBackground(C_WHITE);
            btnExcel.setForeground(C_PRIMARY);
            btnExcel.setFocusPainted(false);
            btnExcel.setBorder(BorderFactory.createCompoundBorder(
                    BorderFactory.createLineBorder(new Color(191, 219, 254), 1, true),
                    BorderFactory.createEmptyBorder(4, 10, 4, 10)));
            btnExcel.setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
            int kIdxFinal = kIdx;
            btnExcel.addActionListener(e -> onExportKelasExcel(kIdxFinal));

            JButton btnCetak = new JButton("Cetak");
            btnCetak.setFont(new Font("Segoe UI", Font.PLAIN, 11));
            btnCetak.setBackground(C_WHITE);
            btnCetak.setForeground(C_SUCCESS);
            btnCetak.setFocusPainted(false);
            btnCetak.setBorder(BorderFactory.createCompoundBorder(
                    BorderFactory.createLineBorder(new Color(167, 243, 208), 1, true),
                    BorderFactory.createEmptyBorder(4, 10, 4, 10)));
            btnCetak.setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
            btnCetak.addActionListener(e -> onPrintKelas(kIdxFinal, namaKelas));

            btnPanel.add(btnExcel);
            btnPanel.add(btnCetak);

            row.add(lbl, BorderLayout.CENTER);
            row.add(btnPanel, BorderLayout.EAST);
            kelasPanel.add(row);
            kelasPanel.add(Box.createVerticalStrut(2));
        }

        // Show the panel
        pnlResult.setVisible(true);
        // Scroll to top
        SwingUtilities.invokeLater(() -> {
            java.awt.Container parent = pnlResult.getParent();
            while (parent != null && !(parent instanceof JScrollPane))
                parent = parent.getParent();
            if (parent instanceof JScrollPane) {
                ((JScrollPane) parent).getVerticalScrollBar().setValue(0);
            }
        });
    }

    private void onExportGuruExcel(String guruNum) {
        try {
            String namaGuru = "";
            for (String[] data : scheduleKebutuhan) {
                String num = data[0].replaceAll("[^0-9]","");
                if (num.equals(guruNum)) { namaGuru = data[1]; break; }
            }
            String fileName = "Guru_" + guruNum + "_" + namaGuru.replaceAll("[^a-zA-Z0-9]", "_") + ".xlsx";
            JFileChooser fc = new JFileChooser();
            fc.setDialogTitle("Simpan Excel Guru");
            fc.setSelectedFile(new File(fileName));
            if (fc.showSaveDialog(this) != JFileChooser.APPROVE_OPTION) return;
            File outFile = fc.getSelectedFile();
            if (outFile == null) return;
            String path = outFile.getAbsolutePath();
            if (!path.endsWith(".xlsx")) path += ".xlsx";
            Main.exportSingleGuruToFile(guruNum, scheduleJadwal, scheduleHariRange, scheduleKebutuhan, path);
            appendLog("  Excel guru " + guruNum + " -> " + outFile.getName(), LOG_OK);
            Desktop.getDesktop().open(new File(path));
        } catch (Exception ex) {
            appendLog("  Gagal export guru " + guruNum + ": " + ex.getMessage(), LOG_ERR);
            JOptionPane.showMessageDialog(this, "Gagal export Excel:\n" + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
        }
    }

    private void onExportKelasExcel(int kelasIdx) {
        try {
            String[] allKelas = Main.buatDaftarKelas();
            String namaKelas = allKelas[kelasIdx];
            String fileName = "Kelas_" + namaKelas + ".xlsx";
            JFileChooser fc = new JFileChooser();
            fc.setDialogTitle("Simpan Excel Kelas");
            fc.setSelectedFile(new File(fileName));
            if (fc.showSaveDialog(this) != JFileChooser.APPROVE_OPTION) return;
            File outFile = fc.getSelectedFile();
            if (outFile == null) return;
            String path = outFile.getAbsolutePath();
            if (!path.endsWith(".xlsx")) path += ".xlsx";
            Main.exportSingleKelasToFile(kelasIdx, scheduleJadwal, scheduleHariRange, scheduleKebutuhan, path);
            appendLog("  Excel kelas " + namaKelas + " -> " + outFile.getName(), LOG_OK);
            Desktop.getDesktop().open(new File(path));
        } catch (Exception ex) {
            appendLog("  Gagal export kelas: " + ex.getMessage(), LOG_ERR);
            JOptionPane.showMessageDialog(this, "Gagal export Excel:\n" + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
        }
    }

    private void onPrintGuru(String guruNum, String namaGuru) {
        Object[] result = Main.generateGuruPreview(guruNum, scheduleJadwal, scheduleHariRange, scheduleKebutuhan);
        String[][] data = (String[][]) result[0];
        String[] headers = (String[]) result[1];
        int totalJam = 0;
        for (int r = 1; r < data.length; r++) {
            String jam = data[r][1];
            String kls = data[r][3];
            if (jam != null && !jam.isEmpty()
                && kls != null && !kls.isEmpty()
                && !kls.contains("Sholat") && !kls.contains("Istirahat")) {
                totalJam++;
            }
        }
        String[] infoLines = {
            "KODE GURU       : " + guruNum,
            "NAMA GURU       : " + namaGuru,
            "JUMLAH MENGAJAR : " + totalJam + " JAM"
        };
        showPrintPreview("Guru " + guruNum + " - " + namaGuru, data, headers, infoLines);
    }

    private void onPrintKelas(int kelasIdx, String namaKelas) {
        Object[] result = Main.generateKelasPreview(kelasIdx, scheduleJadwal, scheduleHariRange, scheduleKebutuhan);
        String[][] data = (String[][]) result[0];
        String[] headers = (String[]) result[1];
        String[] infoLines = { "KELAS : " + namaKelas };
        showPrintPreview("Kelas " + namaKelas, data, headers, infoLines);
    }

    private void showPrintPreview(String title, String[][] data, String[] headers, String[] infoLines) {
        // Save original HARI values before in-place merge
        String[] originalHari = new String[data.length];
        for (int r = 1; r < data.length; r++) {
            originalHari[r] = data[r][0];
        }

        // Merge HARI in-place: hanya tampilkan di baris pertama tiap hari
        String lastHari = data.length > 1 ? data[1][0] : "";
        for (int r = 2; r < data.length; r++) {
            if (data[r][0].equals(lastHari)) {
                data[r][0] = "";
            } else {
                lastHari = data[r][0];
            }
        }

        // Split data by day (using original hari values)
        java.util.LinkedHashMap<String, java.util.List<String[]>> dayData = new java.util.LinkedHashMap<>();
        String curDay = "";
        for (int r = 1; r < data.length; r++) {
            String day = originalHari[r];
            if (!day.isEmpty()) curDay = day;
            String[] rowCopy = data[r].clone();
            dayData.computeIfAbsent(curDay, k -> new java.util.ArrayList<>()).add(rowCopy);
        }

        // Merge HARI within each day's copy
        for (java.util.List<String[]> rows : dayData.values()) {
            if (!rows.isEmpty()) {
                String first = rows.get(0)[0];
                for (int ri = 1; ri < rows.size(); ri++) {
                    rows.get(ri)[0] = "";
                }
            }
        }

        // Active day list (only days with data)
        String[] dayOrder = {"SENIN","SELASA","RABU","KAMIS","JUMAT"};
        java.util.List<String> activeDays = new java.util.ArrayList<>();
        for (String d : dayOrder) {
            if (dayData.containsKey(d) && !dayData.get(d).isEmpty()) activeDays.add(d);
        }
        if (activeDays.isEmpty()) return;

        // --- Dialog ---
        JDialog dialog = new JDialog(this, "Preview - " + title, true);
        dialog.setDefaultCloseOperation(JDialog.DISPOSE_ON_CLOSE);
        dialog.setSize(820, 620);
        dialog.setLocationRelativeTo(this);

        // Title + info header
        JPanel infoPanel = new JPanel();
        infoPanel.setLayout(new BoxLayout(infoPanel, BoxLayout.Y_AXIS));
        infoPanel.setBackground(Color.WHITE);
        infoPanel.setBorder(BorderFactory.createEmptyBorder(6, 16, 0, 16));
        for (String line : infoLines) {
            JLabel lbl = new JLabel(line);
            lbl.setFont(new Font("Segoe UI", Font.BOLD, 13));
            lbl.setForeground(new Color(30, 41, 59));
            infoPanel.add(lbl);
        }

        JLabel lblTitle = new JLabel("Jadwal - " + title);
        lblTitle.setFont(new Font("Segoe UI", Font.BOLD, 16));
        lblTitle.setBorder(BorderFactory.createEmptyBorder(8, 12, 2, 12));
        lblTitle.setForeground(new Color(30, 41, 59));

        // Day buttons
        JPanel dayBtnPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 6, 6));
        dayBtnPanel.setBackground(Color.WHITE);
        dayBtnPanel.setBorder(BorderFactory.createEmptyBorder(2, 12, 4, 12));
        Color btnNormalBg = Color.WHITE;
        Color btnNormalFg = new Color(37, 99, 235);
        Color btnSelBg    = new Color(219, 234, 254);  // light blue
        Color btnSelFg    = new Color(37, 99, 235);    // blue text (always visible)
        java.util.List<JButton> dayBtns = new java.util.ArrayList<>();
        for (String day : activeDays) {
            JButton btn = new JButton(day);
            btn.setFont(new Font("Segoe UI", Font.BOLD, 13));
            btn.setFocusPainted(false);
            btn.setBackground(btnNormalBg);
            btn.setForeground(btnNormalFg);
            btn.setOpaque(true);
            btn.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createLineBorder(new Color(37, 99, 235), 1, true),
                BorderFactory.createEmptyBorder(4, 14, 4, 14)));
            dayBtns.add(btn);
            dayBtnPanel.add(btn);
        }

        // JTable — initially empty
        DefaultTableModel tableModel = new DefaultTableModel(headers, 0) {
            @Override public boolean isCellEditable(int r, int c) { return false; }
        };
        JTable table = new JTable(tableModel);
        table.setFont(new Font("Segoe UI", Font.PLAIN, 12));
        table.getTableHeader().setFont(new Font("Segoe UI", Font.BOLD, 12));
        table.getTableHeader().setBackground(new Color(37, 99, 235));
        table.getTableHeader().setForeground(Color.WHITE);
        table.getTableHeader().setReorderingAllowed(false);
        table.setRowHeight(24);
        table.setShowGrid(true);
        table.setGridColor(new Color(209, 213, 219));
        table.setIntercellSpacing(new Dimension(1, 1));
        table.setAutoResizeMode(JTable.AUTO_RESIZE_SUBSEQUENT_COLUMNS);

        boolean isGuruTbl = headers.length > 3 && "KLS".equals(headers[3]);
        table.getColumnModel().getColumn(0).setPreferredWidth(isGuruTbl ? 60 : 55);
        table.getColumnModel().getColumn(1).setPreferredWidth(30);
        table.getColumnModel().getColumn(2).setPreferredWidth(120);
        table.getColumnModel().getColumn(3).setPreferredWidth(isGuruTbl ? 70 : 140);
        table.getColumnModel().getColumn(4).setPreferredWidth(isGuruTbl ? 220 : 200);

        // Alternating row colors + special row detection
        Color bgLight = new Color(239, 246, 255);
        Color bgCyan  = new Color(219, 234, 254);
        table.setDefaultRenderer(Object.class, new DefaultTableCellRenderer() {
            @Override
            public Component getTableCellRendererComponent(JTable t, Object v, boolean sel, boolean fcs, int r, int c) {
                Component comp = super.getTableCellRendererComponent(t, v, sel, fcs, r, c);
                if (!sel && t.isEnabled()) {
                    String txt = v != null ? v.toString() : "";
                    boolean special = txt.contains("Sholat") || txt.contains("Istirahat");
                    if (special) {
                        comp.setBackground(bgCyan);
                        if (comp instanceof JLabel) ((JLabel) comp).setFont(t.getFont().deriveFont(Font.ITALIC));
                    } else {
                        comp.setBackground(r % 2 == 0 ? bgLight : Color.WHITE);
                        if (comp instanceof JLabel) ((JLabel) comp).setFont(t.getFont());
                    }
                }
                return comp;
            }
        });

        // Populate first day
        String[] finalActiveDays = activeDays.toArray(new String[0]);
        {
            java.util.List<String[]> dayRows = dayData.get(finalActiveDays[0]);
            for (String[] row : dayRows) tableModel.addRow(row);
            dayBtns.get(0).setBackground(btnSelBg);
            dayBtns.get(0).setForeground(btnSelFg);
        }

        // Day button actions
        for (int bi = 0; bi < finalActiveDays.length; bi++) {
            final int idx = bi;
            String day = finalActiveDays[idx];
            JButton btn = dayBtns.get(idx);
            btn.addActionListener(ev -> {
                // Update button highlights
                for (int j = 0; j < finalActiveDays.length; j++) {
                    boolean isThis = j == idx;
                    dayBtns.get(j).setBackground(isThis ? btnSelBg : btnNormalBg);
                    dayBtns.get(j).setForeground(isThis ? btnSelFg : btnNormalFg);
                }
                // Update table
                tableModel.setRowCount(0);
                java.util.List<String[]> rows = dayData.get(day);
                for (String[] row : rows) tableModel.addRow(row);
                SwingUtilities.invokeLater(() -> {
                    if (table.getParent() != null) table.scrollRectToVisible(new Rectangle(0, 0, 1, 1));
                });
            });
        }

        JScrollPane sp = new JScrollPane(table);
        sp.setBorder(BorderFactory.createEmptyBorder(2, 10, 6, 10));

        // --- Print button ---
        JButton btnPrint = new JButton("Print");
        btnPrint.setFont(new Font("Segoe UI", Font.BOLD, 13));
        btnPrint.setBackground(C_PRIMARY);
        btnPrint.setForeground(Color.WHITE);
        btnPrint.setFocusPainted(false);
        btnPrint.setBorderPainted(false);
        btnPrint.setPreferredSize(new Dimension(100, 36));
        btnPrint.addActionListener(e -> {
            try {
                PrinterJob job = PrinterJob.getPrinterJob();
                job.setJobName("Jadwal - " + title);

                // Fix day data for print (each day = one section with day header + table)
                final String[] printDays = finalActiveDays;
                final java.util.Map<String, java.util.List<String[]>> printDayData = dayData;

                job.setPrintable((graphics, pageFormat, pageIndex) -> {
                    Graphics2D g = (Graphics2D) graphics;
                    g.translate(pageFormat.getImageableX(), pageFormat.getImageableY());

                    boolean isGuru = headers.length > 3 && "KLS".equals(headers[3]);
                    int[] colRatios = isGuru
                        ? new int[]{80, 50, 160, 100, 300}
                        : new int[]{80, 50, 160, 250, 250};
                    int ratioSum = 0;
                    for (int v : colRatios) ratioSum += v;
                    int totalW = (int) pageFormat.getImageableWidth() - 10;
                    int[] colW = new int[colRatios.length];
                    for (int i = 0; i < colRatios.length; i++)
                        colW[i] = totalW * colRatios[i] / ratioSum;

                    int x0 = 4;
                    int padX = 6;
                    int rowH = 18;
                    int dayHdrH = 22;
                    int spacer = 4;

                    g.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON);

                    // Header area height (info + table header, no title in print)
                    g.setFont(new Font("Segoe UI", Font.BOLD, 10));
                    int infoH = g.getFontMetrics().getHeight() * infoLines.length + 6;
                    int hdrH = g.getFontMetrics().getHeight() + 8;
                    int headerAreaH = infoH + hdrH;
                    int availH = (int) pageFormat.getImageableHeight();

                    // Build day section heights: day header + table header + rows
                    int[] dayHeights = new int[printDays.length];
                    for (int d = 0; d < printDays.length; d++) {
                        int rc = printDayData.get(printDays[d]).size();
                        dayHeights[d] = dayHdrH + hdrH + rowH * rc + spacer;
                    }

                    // Allocate pages (group days that fit after header)
                    java.util.List<int[]> pageAlloc = new java.util.ArrayList<>();
                    int di = 0;
                    while (di < printDays.length) {
                        int startDi = di;
                        int usedH = 0;
                        boolean firstDay = true;
                        while (di < printDays.length) {
                            int need = dayHeights[di];
                            if (usedH + need <= availH - headerAreaH) {
                                usedH += need;
                                firstDay = false;
                                di++;
                            } else {
                                break;
                            }
                        }
                        if (firstDay && di < printDays.length) {
                            // Force at least one day per page
                            usedH = dayHeights[di];
                            di++;
                        }
                        pageAlloc.add(new int[]{startDi, di - 1});
                    }
                    if (pageIndex >= pageAlloc.size()) return Printable.NO_SUCH_PAGE;
                    int[] pd = pageAlloc.get(pageIndex);

                    // --- Draw info lines + table header (repeated on every page) ---
                    int ry = 0;

                    g.setFont(new Font("Segoe UI", Font.BOLD, 10));
                    FontMetrics ifm = g.getFontMetrics();
                    g.setColor(new Color(30, 41, 59));
                    for (String line : infoLines) {
                        g.drawString(line, x0, ry + ifm.getAscent());
                        ry += ifm.getHeight();
                    }
                    ry += 6;

                    // --- Each day on this page ---
                    for (int d = pd[0]; d <= pd[1]; d++) {
                        String dayName = printDays[d];
                        java.util.List<String[]> dayRows = printDayData.get(dayName);
                        if (dayRows == null || dayRows.isEmpty()) continue;

                        // Day header bar
                        g.setColor(new Color(209, 213, 219));
                        g.fillRect(x0, ry, totalW, dayHdrH);
                        g.setFont(new Font("Segoe UI", Font.BOLD, 11));
                        FontMetrics dhfm = g.getFontMetrics();
                        int dhTextX = x0 + (totalW - dhfm.stringWidth(dayName)) / 2;
                        int dhTextY = ry + (dayHdrH - dhfm.getHeight()) / 2 + dhfm.getAscent();
                        g.setColor(new Color(30, 41, 59));
                        g.drawString(dayName, dhTextX, dhTextY);
                        ry += dayHdrH;

                        // Table header
                        g.setFont(new Font("Segoe UI", Font.BOLD, 10));
                        FontMetrics hfm = g.getFontMetrics();
                        g.setColor(new Color(209, 213, 219));
                        g.fillRect(x0, ry, totalW, hdrH);
                        g.setColor(Color.BLACK);
                        int cx = x0 + padX;
                        for (int c = 0; c < headers.length; c++) {
                            g.drawString(headers[c], cx, ry + (hdrH - hfm.getHeight()) / 2 + hfm.getAscent());
                            cx += colW[c];
                        }
                        int tableHeaderBottom = ry + hdrH;
                        ry += hdrH;

                        // Data rows
                        g.setFont(new Font("Segoe UI", Font.PLAIN, 9));
                        FontMetrics dfm = g.getFontMetrics();
                        int dataStartRow = ry;
                        for (String[] row : dayRows) {
                            boolean isSpecial = false;
                            for (String v : row) {
                                if (v != null && (v.contains("Sholat") || v.contains("Istirahat"))) {
                                    isSpecial = true;
                                    break;
                                }
                            }

                            if (isSpecial)
                                g.setColor(new Color(219, 234, 254));
                            else if ((d + (dayRows.indexOf(row))) % 2 == 0)
                                g.setColor(new Color(239, 246, 255));
                            else
                                g.setColor(Color.WHITE);
                            g.fillRect(x0, ry, totalW, rowH);

                            if (isSpecial)
                                g.setFont(new Font("Segoe UI", Font.ITALIC, 9));
                            g.setColor(Color.BLACK);
                            cx = x0 + padX;
                            for (int c = 0; c < row.length && c < colW.length; c++) {
                                String v = row[c] != null ? row[c] : "";
                                int textY = ry + (rowH - dfm.getHeight()) / 2 + dfm.getAscent();
                                g.drawString(v, cx, textY);
                                cx += colW[c];
                            }
                            if (isSpecial)
                                g.setFont(new Font("Segoe UI", Font.PLAIN, 9));

                            ry += rowH;
                        }
                        int dataEndRow = ry;

                        // Grid lines for this day
                        int gridH = dataEndRow - dataStartRow;
                        if (gridH > 0) {
                            g.setStroke(new BasicStroke(2f));
                            g.setColor(new Color(156, 163, 175));
                            // Horizontal
                            for (int gr = 0; gr <= (gridH / rowH); gr++) {
                                int ly = dataStartRow + gr * rowH;
                                g.drawLine(x0, ly, x0 + totalW, ly);
                            }
                            // Vertical
                            cx = x0;
                            for (int c = 0; c <= colW.length; c++) {
                                g.drawLine(cx, dataStartRow, cx, dataEndRow);
                                if (c < colW.length) cx += colW[c];
                            }
                            g.setStroke(new BasicStroke(1f));
                        }

                        ry += spacer;
                    }

                    return Printable.PAGE_EXISTS;
                });

                if (job.printDialog()) {
                    job.print();
                    appendLog("  Cetak " + title + " dikirim ke printer", LOG_OK);
                }
            } catch (Exception ex) {
                appendLog("  Gagal cetak: " + ex.getMessage(), LOG_ERR);
                JOptionPane.showMessageDialog(dialog, "Gagal mencetak:\n" + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
            }
        });

        // --- Layout ---
        JPanel topPanel = new JPanel(new BorderLayout());
        JPanel titleInfo = new JPanel(new BorderLayout());
        titleInfo.setBackground(Color.WHITE);
        titleInfo.add(lblTitle, BorderLayout.NORTH);
        titleInfo.add(infoPanel, BorderLayout.CENTER);

        JPanel centerArea = new JPanel(new BorderLayout());
        centerArea.add(dayBtnPanel, BorderLayout.NORTH);
        centerArea.add(sp, BorderLayout.CENTER);

        topPanel.add(titleInfo, BorderLayout.NORTH);
        topPanel.add(centerArea, BorderLayout.CENTER);

        JPanel bottom = new JPanel(new FlowLayout(FlowLayout.CENTER, 8, 8));
        bottom.add(btnPrint);
        JButton btnClose = new JButton("Tutup");
        btnClose.setFont(new Font("Segoe UI", Font.PLAIN, 12));
        btnClose.addActionListener(ev -> dialog.dispose());
        bottom.add(btnClose);

        dialog.setLayout(new BorderLayout());
        dialog.add(topPanel, BorderLayout.CENTER);
        dialog.add(bottom, BorderLayout.SOUTH);
        dialog.setVisible(true);
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