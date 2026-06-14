package org.example;
import org. apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel. XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

import javax.swing.text.LayeredHighlighter;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Set;
import java.util.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFFont;

public class Main {

    public static void main(String[] args) throws Exception {

        SchedulerUI.launch(Main::runScheduler);
    }
    static void runScheduler() throws Exception {
        iterasiGlobal[0] = 0;
        FileInputStream fis = new FileInputStream(SchedulerUI.inputFilePath);
        Workbook wb = new XSSFWorkbook(fis);
        Sheet sheet = wb.getSheetAt(0);

        //arraylist untuk menyimpan kebutuhan tugas mengajar guru
        ArrayList<String[]> kebutuhan = new ArrayList<>();

        Row headerKelas = sheet.getRow(5);
        int totalBaris = sheet.getPhysicalNumberOfRows();

        String namaGuru = "";
        String nomorGuru = "";
        String mapel = "";

        // Baca data dari Excel
        for (int r = 0; r < totalBaris; r++) {

            Row row = sheet.getRow(r);
            if (row == null) continue;

            String kolomNo = getString(row.getCell(0));
            String kolomNama = getString(row.getCell(1));
            String kolomMapel = getString(row.getCell(2));

            // Update nama guru dan nomor
            if (!kolomNama.equals("") && !kolomNama.toLowerCase().contains("nip")) {
                namaGuru = kolomNama;
                nomorGuru = kolomNo;
            }

            // Update mapel
            if (!kolomMapel.equals("")) {
                mapel = kolomMapel;
            }

            if (namaGuru.equals("") || mapel.equals("") || nomorGuru.equals("")) continue;

            // Loop kelas
            for (int c = 3; ; c++) {

                Cell headerCell = headerKelas.getCell(c);
                String headerText = getString(headerCell);

                if (headerText.equalsIgnoreCase("JML") || headerText.equals("")) {
                    break;
                }

                String beban = getString(row.getCell(c));

                if (isAngka(beban) && !beban.equals("0")) {

                    int indexKelas = c - 3;
                    String namaKelas = String.valueOf(indexKelas);

                    String[] data = new String[5];
                    data[0] = nomorGuru;  // Sementara nomor aja
                    data[1] = namaGuru;
                    data[2] = mapel;
                    data[3] = namaKelas;
                    data[4] = beban;

                    kebutuhan.add(data);
                }
            }
        }

        // Memperbaiki ID berdasarkan jumlah mapel
        for (int i = 0; i < kebutuhan.size(); i++) {
            String[] current = kebutuhan.get(i);
            String nomor = current[0];
            String guru = current[1];
            String mapelCurrent = current[2];

            // Cari semua mapel unik untuk guru ini
            ArrayList<String> daftarMapel = new ArrayList<>();
            for (String[] data : kebutuhan) {
                if (data[1].equals(guru)) {  // Sama guru
                    if (!daftarMapel.contains(data[2])) {  // Mapel belum ada di list
                        daftarMapel.add(data[2]);
                    }
                }
            }

            // Bikin ID
            if (daftarMapel.size() == 1) {
                // Cuma 1 mapel == nomor aja
                current[0] = nomor;
            } else {
                // Lebih dari 1 mapel menambahkan suffix
                int urutanMapel = daftarMapel.indexOf(mapelCurrent);
                char suffix = (char) ('a' + urutanMapel);
                current[0] = nomor + suffix;
            }
        }

        wb.close();

//        // print data guru di arraylist kebutuhan
//        System.out.println("=== DATA GURU DAN BEBAN ===");
//        for (String[] d : kebutuhan) {
//            System.out.println(d[0] + " | " + d[1] + " | " + d[2] + " | " + d[3] + " | " + d[4]);
//        }
//        System.out.println("Total data: " + kebutuhan.size());

        // menyimpan id gutu PJOK
        Set<String> guruPJOK = new HashSet<>();
        for (int i = 0; i < kebutuhan.size(); i++) {

            String[] data = kebutuhan.get(i);

            String idGuru = data[0];
            mapel = data[2];

            if (mapel.equalsIgnoreCase("PJOK")) {
                guruPJOK.add(idGuru);
            }
        }

        // Menyimpan id guru Matematika
        Set<String> guruMatematika = new HashSet<>();
        for (int i = 0; i < kebutuhan.size(); i++) {

            String[] data = kebutuhan.get(i);

            String idGuru = data[0];
            mapel = data[2];

            if (mapel.equalsIgnoreCase("MATEMATIKA")) {
                guruMatematika.add(idGuru);
            }
        }

        // MGMP Senin (PKN, PJOK, IPS)
        Set<String> MGMPsenin = new HashSet<>();

        for (int i = 0; i < kebutuhan.size(); i++) {

            String[] data = kebutuhan.get(i);

            String idGuruAsli = data[0];
            mapel = data[2];

            if (mapel.equalsIgnoreCase("PKN") ||
                    mapel.equalsIgnoreCase("PJOK") ||
                    mapel.equalsIgnoreCase("IPS")) {

                // cek apakah ada huruf
                if (idGuruAsli.matches(".*[a-zA-Z].*")) {

                    // ambil hanya yang ada huruf A/a
                    if (idGuruAsli.toLowerCase().contains("a")) {
                        String idGuru = idGuruAsli.replaceAll("[^0-9]", "");
                        MGMPsenin.add(idGuru);
                    }

                } else {
                    // kalau pure angka langsung ambil
                    MGMPsenin.add(idGuruAsli);
                }
            }
        }

        // MGMP Selasa (B. Inggris, B. Indonesia, B. Jawa)
        Set<String> MGMPselasa = new HashSet<>();

        for (int i = 0; i < kebutuhan.size(); i++) {

            String[] data = kebutuhan.get(i);

            String idGuruAsli = data[0];
            mapel = data[2];

            if (mapel.equalsIgnoreCase("B. INGGRIS") ||
                    mapel.equalsIgnoreCase("B. INDONESIA") ||
                    mapel.equalsIgnoreCase("B. JAWA")) {

                // cek apakah ada huruf
                if (idGuruAsli.matches(".*[a-zA-Z].*")) {

                    // ambil hanya yang ada huruf A/a
                    if (idGuruAsli.toLowerCase().contains("a")) {
                        String idGuru = idGuruAsli.replaceAll("[^0-9]", "");
                        MGMPselasa.add(idGuru);
                    }

                } else {
                    // kalau pure angka langsung ambil
                    MGMPselasa.add(idGuruAsli);
                }
            }
        }


        Set<String> MGMPrabu = new HashSet<>();

        for (int i = 0; i < kebutuhan.size(); i++) {

            String[] data = kebutuhan.get(i);

            String idGuruAsli = data[0];
            mapel = data[2];

            if (mapel.equalsIgnoreCase("SKI") ||
                    mapel.equalsIgnoreCase("AQIDAH A.") ||
                    mapel.equalsIgnoreCase("B. ARAB") ||
                    mapel.equalsIgnoreCase("QURDITS") ||
                    mapel.equalsIgnoreCase("FIQIH")) {

                // cek apakah ada huruf
                if (idGuruAsli.matches(".*[a-zA-Z].*")) {

                    // ambil hanya yang ada huruf A/a
                    if (idGuruAsli.toLowerCase().contains("a")) {
                        String idGuru = idGuruAsli.replaceAll("[^0-9]", "");
                        MGMPrabu.add(idGuru);
                    }

                } else {
                    // kalau pure angka langsung ambil
                    MGMPrabu.add(idGuruAsli);
                }
            }
        }

        // MGMP Kamis (MATEMATIKA, IPA, BK, INFORMATIKA)
        Set<String> MGMPkamis = new HashSet<>();

        for (int i = 0; i < kebutuhan.size(); i++) {

            String[] data = kebutuhan.get(i);

            String idGuruAsli = data[0];
            mapel = data[2];

            if (mapel.equalsIgnoreCase("MATEMATIKA") ||
                    mapel.equalsIgnoreCase("IPA") ||
                    mapel.equalsIgnoreCase("BK") ||
                    mapel.equalsIgnoreCase("INFORMATIKA")) {

                // cek apakah ada huruf
                if (idGuruAsli.matches(".*[a-zA-Z].*")) {

                    // ambil hanya yang ada huruf A/a
                    if (idGuruAsli.toLowerCase().contains("a")) {
                        String idGuru = idGuruAsli.replaceAll("[^0-9]", "");
                        MGMPkamis.add(idGuru);
                    }

                } else {
                    // kalau pure angka langsung ambil
                    MGMPkamis.add(idGuruAsli);
                }
            }
        }


        int jumlahKelas = SchedulerUI.jumlahKelas;
        int senin       = SchedulerUI.jamSenin;
        int selasa      = SchedulerUI.jamSelasa;
        int rabu        = SchedulerUI.jamRabu;
        int kamis       = SchedulerUI.jamKamis;
        int jumat       = SchedulerUI.jamJumat;

        int totalJam = senin + selasa + rabu + kamis + jumat;

        String[][] jadwal = new String[totalJam][jumlahKelas];

        for (int i = 0; i < totalJam; i++) {
            for (int j = 0; j < jumlahKelas; j++) {
                jadwal[i][j] = "";
            }
        }

        //jadwal fleksibel
        int[][] hariRange = new int[5][2];

        int start = 0;

        hariRange[0][0] = start;
        hariRange[0][1] = start + senin - 1;
        start += senin;

        hariRange[1][0] = start;
        hariRange[1][1] = start + selasa - 1;
        start += selasa;

        hariRange[2][0] = start;
        hariRange[2][1] = start + rabu - 1;
        start += rabu;

        hariRange[3][0] = start;
        hariRange[3][1] = start + kamis - 1;
        start += kamis;

        hariRange[4][0] = start;
        hariRange[4][1] = start + jumat - 1;


        System.out.println("=== Daftar Tugas Mengajar Guru ===");
        String[] allKelas = buatDaftarKelas();
        for (String[] d : kebutuhan) {
            int kIdx = Integer.parseInt(d[3]);
            String namaKelas = (kIdx >= 0 && kIdx < allKelas.length) ? allKelas[kIdx] : d[3];
            System.out.println(d[0] + " | " + d[1] + " | " + d[2] + " | " + namaKelas + " | " + d[4]);
        }
        System.out.println("Total data: " + kebutuhan.size());
        int totalBeban = 0;
        for (String[] d : kebutuhan) {
            totalBeban += Integer.parseInt(d[4]);
        }
        System.out.println("Total beban awal: " + totalBeban);

        System.out.println();
        System.out.println("Memulai Initial Solution.....");

        //Memulai Initial Solution
        initialSolutionPJOK(kebutuhan, jadwal, hariRange);

        kebutuhan.sort((a, b) -> Integer.parseInt(a[4]) - Integer.parseInt(b[4]));
       //kebutuhan.sort((a, b) -> Integer.parseInt(b[4]) - Integer.parseInt(a[4]));
        //initialSolutionBlok3(kebutuhan, jadwal, hariRange);
        initialSolutionBlok2(kebutuhan, jadwal, hariRange);
        initialSolutionSwap2(kebutuhan, jadwal, hariRange, guruPJOK);
        initialSolutionBlok1(kebutuhan, jadwal, hariRange);
        System.out.println("Initial Solution Berhasil");

        System.out.println();
        System.out.println("=== SISA BEBAN SETELAH INITIAL SOLUTION ===");
        int totalSisa = 0;
        for (String[] d : kebutuhan) {
            int sisa = Integer.parseInt(d[4]);
            totalSisa += sisa;
            int kIdx = Integer.parseInt(d[3]);
            String namaKelas = (kIdx >= 0 && kIdx < allKelas.length) ? allKelas[kIdx] : d[3];
            if (sisa > 0) {
                System.out.println("[BELUM] " + d[0] + " | " + d[1] + " | " + d[2] + " | " + namaKelas + " | sisa: " + sisa);
            } else {
                System.out.println("[OK]    " + d[0] + " | " + d[1] + " | " + d[2] + " | " + namaKelas + " | sisa: 0");
            }
        }
        System.out.println("Total sisa beban: " + totalSisa);
        System.out.println("Total data: " + kebutuhan.size());

        long mulai = System.nanoTime();

        System.out.println();
        System.out.println("Memulai optimasi....");
        System.out.println("Memulai optimasi ID guru 43....");
        jadwal = hillClimbingTabuGuru43(jadwal, hariRange, guruPJOK);
        System.out.println();
        System.out.println("Memulai optimasi Hard Constraint....");
        jadwal = hillClimbingTabuHardConstrain(jadwal, hariRange, guruPJOK);
        System.out.println();
        System.out.println("Memulai optimasi Soft Constraint....");
        jadwal =  LAHCTabuSoftConstrain(jadwal, hariRange, guruPJOK, MGMPsenin, MGMPselasa, MGMPrabu, MGMPkamis, guruMatematika);


        //HillClimbingRandom
       //jadwal = hillClimbingRandomHardConstrain(jadwal, hariRange, guruPJOK);
       //jadwal = hillClimbingRandomSoftConstrain(jadwal, hariRange, guruPJOK, MGMPsenin, MGMPselasa, MGMPrabu, MGMPkamis, guruMatematika);

       //HillClimbingTabu
//        jadwal = hillClimbingTabuGuru43(jadwal, hariRange, guruPJOK);
//        jadwal = hillClimbingTabuHardConstrain(jadwal, hariRange, guruPJOK);
        //jadwal = hillClimbingTabuSoftConstrain(jadwal, hariRange, guruPJOK, MGMPsenin, MGMPselasa, MGMPrabu, MGMPkamis, guruMatematika);

        // OPTIMASI FIX
        //jadwal = hillClimbingTabuGuru43(jadwal, hariRange, guruPJOK);
        //jadwal = hillClimbingTabuHardConstrain(jadwal, hariRange, guruPJOK);
        //jadwal =  LAHCTabuSoftConstrain(jadwal, hariRange, guruPJOK, MGMPsenin, MGMPselasa, MGMPrabu, MGMPkamis, guruMatematika);

        System.out.println();
        System.out.println("=== Daftar Pelanggaran ===");
        System.out.println("Pelanggaran PJOK: " + hitungPenaltiPJOK(jadwal, hariRange, guruPJOK));
        System.out.println("Pelanggaran guru ID 43: " + hitungPenaltiGuruID43(jadwal, hariRange));
        System.out.println("Pelanggaran MGMP Senin: " + hitungPenaltiMGMPSenin(jadwal, hariRange, MGMPsenin));
        System.out.println("Pelanggaran MGMP Selasa: " + hitungPenaltiMGMPSelasa(jadwal, hariRange, MGMPselasa));
        System.out.println("Pelanggaran MGMP Rabu: " + hitungPenaltiMGMPRabu(jadwal, hariRange, MGMPrabu));
        System.out.println("Pelanggaran MGMP Kamis: " + hitungPenaltiMGMPKamis(jadwal, hariRange, MGMPkamis));
        System.out.println("Pelanggaran MGMP awal total: " + hitungTotalPenaltiMGMP(jadwal, hariRange, MGMPsenin, MGMPselasa, MGMPrabu, MGMPkamis));
        System.out.println("Pelanggaran Matematika: " + hitungPenaltiMatematika(jadwal, hariRange, guruMatematika));
        System.out.println("Pelanggaran jam 9&10: " + hitungPenaltiJam9dan10(jadwal, hariRange, 5));
        System.out.println("Pelanggaran bentrok: " + hitungBentrok(jadwal));
        System.out.println("Pelanggaran max jam per hari: " + hitungPenaltiMaxJamPerHari(jadwal, hariRange, 8, 2));
        System.out.println("Pelanggaran total awal: " + hitungTotalSemuaPenalti(jadwal,
                hariRange,
                guruPJOK,
                MGMPsenin,
                MGMPselasa,
                MGMPrabu,
                MGMPkamis,
                guruMatematika))
        ;
        long end = System.nanoTime();

        double durasiDetik = (end - mulai) / 1_000_000_000.0;
        if (durasiDetik < 60) {
            System.out.println("Waktu eksekusi: " + String.format("%.2f", durasiDetik) + " detik");
        } else {
            double durasiMenit = durasiDetik / 60.0;
            System.out.println("Waktu eksekusi: " + String.format("%.2f", durasiMenit) + " menit");
        }
        exportJadwalToExcel(jadwal, hariRange, SchedulerUI.outputFilePath, kebutuhan);
        SchedulerUI.setData(jadwal, hariRange, kebutuhan);
        SchedulerUI.markDone(true, null);


    }

    // Method
    static String getString(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                return String.valueOf((int) cell.getNumericCellValue());
        }
        return "";
    }

    static boolean isAngka(String s) {
        return s != null && s.matches("\\d+");
    }

    // get hari fleksibel
    static int getHari(int row, int[][] hariRange) {

        for (int i = 0; i < hariRange.length; i++) {

            if (row >= hariRange[i][0] && row <= hariRange[i][1]) {
                return i;
            }
        }
        return -1;
    }


    // Bagianblok Fleksibel
    static boolean bagianBlok(String[][] jadwal, int index, int kelas, String guru, int[][] hariRange) {

        // cek sebelumnya
        if (index - 1 >= 0 &&
                getHari(index - 1, hariRange) == getHari(index, hariRange) &&
                jadwal[index - 1][kelas].equals(guru)) {

            return false;
        }

        // cek sesudahnya
        if (index + 1 < jadwal.length &&
                getHari(index + 1, hariRange) == getHari(index, hariRange) &&
                jadwal[index + 1][kelas].equals(guru)) {

            return false;
        }

        return true;
    }

    static boolean guruSudahAdaDiHari(String[][] jadwal,
                                      int kelas,
                                      String idGuru,
                                      int hari,
                                      int[][] hariRange) {

        int start = hariRange[hari][0];
        int end = hariRange[hari][1];

        for (int r = start; r <= end; r++) {
            if (jadwal[r][kelas].equals(idGuru)) {
                return true;
            }
        }
        return false;
    }

    //Method Initial Solution
    private static void initialSolutionPJOK(List<String[]> kebutuhan, String[][] jadwal, int[][] hariRange) {

        for (int i = 0; i < kebutuhan.size(); i++) {

            String[] data = kebutuhan.get(i);

            String idGuru = data[0];
            String mapel  = data[2];
            int kelas     = Integer.parseInt(data[3]);
            int beban     = Integer.parseInt(data[4]);

            if (!mapel.equalsIgnoreCase("PJOK")) continue;

            int ukuranBlok = 2;

            while (beban >= ukuranBlok) {

                boolean berhasil = false;

                outer:
                for (int h = 0; h < hariRange.length; h++) {

                    int start      = hariRange[h][0];
                    int end        = hariRange[h][1];
                    int batasAkhir = Math.min(start + 4, end);

                    for (int r = start; r <= batasAkhir - 1; r++) {

                        // Cek slot kelas kosong
                        boolean slotKosong = true;
                        for (int cek = r; cek < r + ukuranBlok; cek++) {
                            if (!jadwal[cek][kelas].equals("")) { slotKosong = false; break; }
                        }
                        if (!slotKosong) continue;

                        // Cek guru tidak bentrok lintas kelas
                        boolean bentrokGuru = false;
                        for (int cek = r; cek < r + ukuranBlok && !bentrokGuru; cek++) {
                            for (int k = 0; k < jadwal[0].length; k++) {
                                if (k == kelas) continue;
                                String isi = jadwal[cek][k];
                                if (isi != null && isi.equals(idGuru)) { bentrokGuru = true; break; }
                            }
                        }
                        if (bentrokGuru) continue;

                        // Cek guru tidak muncul dua kali di hari yang sama
                        if (guruSudahAdaDiHari(jadwal, kelas, idGuru, h, hariRange)) continue;

                        // Masukkan blok ke jadwal
                        for (int isi = r; isi < r + ukuranBlok; isi++) jadwal[isi][kelas] = idGuru;

                        beban -= ukuranBlok;
                        data[4] = String.valueOf(beban);
                        berhasil = true;
                        break outer;
                    }
                }

                if (!berhasil) break;
            }
        }
    }

    private static void initialSolutionBlok3(List<String[]> kebutuhan, String[][] jadwal, int[][] hariRange) {

        boolean masihAda = true;

        while (masihAda) {

            masihAda = false;

            for (int i = 0; i < kebutuhan.size(); i+=3) {

                String[] data = kebutuhan.get(i);

                String idGuru = data[0];
                String mapel  = data[2];
                int kelas     = Integer.parseInt(data[3]);
                int beban     = Integer.parseInt(data[4]);

                if (beban != 3 && beban != 5) continue;
                //if (!mapel.equalsIgnoreCase("MATEMATIKA") && !mapel.equalsIgnoreCase("IPA")) continue ;


                int ukuranBlok = 3;

                while (beban >= ukuranBlok) {

                    boolean berhasil = false;

                    for (int h = 0; h < hariRange.length; h++) {

                        int start = hariRange[h][0];
                        int end   = hariRange[h][1];

                        if (guruSudahAdaDiHari(jadwal, kelas, idGuru, h, hariRange)) continue;

                        int kosongBerurutan = 0;

                        for (int r = start; r <= end; r++) {

                            if (jadwal[r][kelas].equals("")) {
                                kosongBerurutan++;
                            } else {
                                kosongBerurutan = 0;
                            }

                            if (kosongBerurutan == ukuranBlok) {

                                for (int isi = r - ukuranBlok + 1; isi <= r; isi++) {
                                    jadwal[isi][kelas] = idGuru;
                                }

                                beban -= ukuranBlok;
                                data[4] = String.valueOf(beban);

                                berhasil = true;
                                masihAda = true;
                                break;
                            }
                        }

                        if (berhasil) break;
                    }

                    if (!berhasil) break;

                    // beban 5 → ambil 3 → sisa 2 → stop
                    if (beban == 2) break;
                }
            }
        }
    }

    private static void initialSolutionBlok2(List<String[]> kebutuhan, String[][] jadwal, int[][] hariRange) {

        boolean masihAda = true;

        while (masihAda) {

            masihAda = false;

            for (int i = 0; i < kebutuhan.size(); i++) {

                String[] data = kebutuhan.get(i);

                String idGuru = data[0];
                int kelas     = Integer.parseInt(data[3]);
                int beban     = Integer.parseInt(data[4]);

                if (beban == 3) continue;

                int ukuranBlok = 2;

                while (beban >= ukuranBlok) {

                    boolean berhasil = false;

                    for (int h = 0; h < hariRange.length; h++) {

                        int hariStart = hariRange[h][0];
                        int end       = hariRange[h][1];

                        if (guruSudahAdaDiHari(jadwal, kelas, idGuru, h, hariRange)) continue;

                        int kosongBerurutan = 0;

                        for (int r = hariStart; r <= end; r++) {

                            if (jadwal[r][kelas].equals("")) {
                                kosongBerurutan++;
                            } else {
                                kosongBerurutan = 0;
                            }

                            if (kosongBerurutan == ukuranBlok) {

                                int rMulai = r - ukuranBlok + 1;

                                if ((rMulai - hariStart) % 2 != 0) continue;

                                for (int isi = rMulai; isi <= r; isi++) {
                                    jadwal[isi][kelas] = idGuru;
                                }

                                beban -= ukuranBlok;
                                data[4] = String.valueOf(beban);

                                berhasil = true;
                                masihAda = true;
                                break;
                            }
                        }

                        if (berhasil) break;
                    }

                    if (!berhasil) break;

                    if (beban == 3) break;
                }
            }
        }
    }

    private static void initialSolutionSwap2(List<String[]> kebutuhan, String[][] jadwal, int[][] hariRange, Set<String> guruPJOK) {

        for (int i = 0; i < kebutuhan.size(); i++) {

            String[] data = kebutuhan.get(i);
            int sisa = Integer.parseInt(data[4]);

            if (sisa < 2) continue;

            String idGuru = data[0];
            int kelas = Integer.parseInt(data[3]);

            while (sisa >= 2) {

                boolean selesai = false;

                for (int r = 0; r < jadwal.length - 1 && !selesai; r++) {

                    if (getHari(r, hariRange) != getHari(r + 1, hariRange)) continue;

                    int hariR = getHari(r, hariRange);
                    int startHari = hariRange[hariR][0];
                    int posisiDalamHari = r - startHari;
                    if (posisiDalamHari % 2 != 0) continue;

                    if (jadwal[r][kelas].equals("") && jadwal[r + 1][kelas].equals("")) {

                        int hariTarget = getHari(r, hariRange);

                        if (!guruSudahAdaDiHari(jadwal, kelas, idGuru, hariTarget, hariRange)) {

                            jadwal[r][kelas] = idGuru;
                            jadwal[r + 1][kelas] = idGuru;

                            sisa -= 2;
                            data[4] = String.valueOf(sisa);

                            selesai = true;
                            break;
                        }

                        if (guruPJOK.contains(idGuru)) break;

                        for (int s = 0; s < jadwal.length - 1 && !selesai; s++) {

                            if (getHari(s, hariRange) != getHari(s + 1, hariRange)) continue;

                            if (getHari(s, hariRange) == 0 && s == hariRange[0][0]) continue;

                            if (!jadwal[s][kelas].equals("") &&
                                    jadwal[s][kelas].equals(jadwal[s + 1][kelas])) {

                                String guruLain = jadwal[s][kelas];

                                if (guruPJOK.contains(guruLain)) continue;

                                boolean bagianDariTiga = false;

                                if (s - 1 >= 0 &&
                                        getHari(s - 1, hariRange) == getHari(s, hariRange) &&
                                        jadwal[s - 1][kelas].equals(guruLain)) {
                                    bagianDariTiga = true;
                                }
                                if (s + 2 < jadwal.length &&
                                        getHari(s + 2, hariRange) == getHari(s, hariRange) &&
                                        jadwal[s + 2][kelas].equals(guruLain)) {
                                    bagianDariTiga = true;
                                }
                                if (bagianDariTiga) continue;

                                int hariAsal = getHari(s, hariRange);
                                if (guruSudahAdaDiHari(jadwal, kelas, guruLain, hariTarget, hariRange)) continue;

                                jadwal[s][kelas] = "";
                                jadwal[s + 1][kelas] = "";

                                jadwal[r][kelas] = guruLain;
                                jadwal[r + 1][kelas] = guruLain;

                                if (!guruSudahAdaDiHari(jadwal, kelas, idGuru, hariAsal, hariRange)) {

                                    jadwal[s][kelas] = idGuru;
                                    jadwal[s + 1][kelas] = idGuru;

                                    sisa -= 2;
                                    data[4] = String.valueOf(sisa);

                                    selesai = true;
                                    break;
                                }

                                // undo kalau gagal
                                jadwal[r][kelas] = "";
                                jadwal[r + 1][kelas] = "";
                                jadwal[s][kelas] = guruLain;
                                jadwal[s + 1][kelas] = guruLain;
                            }
                        }
                    }
                }

                if (!selesai) break;
            }
        }
    }

    private static void initialSolutionBlok1(List<String[]> kebutuhan, String[][] jadwal, int[][] hariRange) {

        for (int i = 0; i < kebutuhan.size(); i++) {

            String[] data = kebutuhan.get(i);
            int sisa = Integer.parseInt(data[4]);

            if (sisa != 1) continue;

            String idGuru = data[0];
            int kelas = Integer.parseInt(data[3]);

            boolean selesai = false;

            for (int r = 0; r < jadwal.length && !selesai; r++) {

                if (!jadwal[r][kelas].equals("")) continue;

                int hariTarget = getHari(r, hariRange);

                if (!guruSudahAdaDiHari(jadwal, kelas, idGuru, hariTarget, hariRange)) {

                    jadwal[r][kelas] = idGuru;
                    data[4] = "0";
                    selesai = true;
                    break;
                }

                for (int s = 0; s < jadwal.length && !selesai; s++) {

                    if (jadwal[s][kelas].equals("")) continue;

                    String guruLain = jadwal[s][kelas];

                    if (guruLain.equals(idGuru)) continue;

                    int hariAsal = getHari(s, hariRange);

                    if (!bagianBlok(jadwal, s, kelas, guruLain, hariRange)) continue;

                    if (guruSudahAdaDiHari(jadwal, kelas, guruLain, hariTarget, hariRange)) continue;

                    jadwal[r][kelas] = guruLain;
                    jadwal[s][kelas] = "";

                    if (!guruSudahAdaDiHari(jadwal, kelas, idGuru, hariAsal, hariRange)) {

                        jadwal[s][kelas] = idGuru;
                        data[4] = "0";
                        selesai = true;
                        break;
                    }

                    // undo
                    jadwal[r][kelas] = "";
                    jadwal[s][kelas] = guruLain;
                }
            }
        }
    }


    static void exportJadwalToExcel(String[][] jadwal, int[][] hariRange, String outputFile, List<String[]> kebutuhan) throws Exception {

        try (Workbook wb = new XSSFWorkbook()) {

            // ----- Warna background -----
            final String GREEN   = "92D050"; // header kelas 7
            final String YELLOW  = "FFFF00"; // header kelas 8
            final String RED     = "FF0000"; // header kelas 9
            final String ORANGE  = "FFC000"; // highlight nilai bentrok di kolom guru
            final String LBLUE   = "BDD7EE"; // header KD/NAMA GURU/JML/HARI
            final String WHITE   = "FFFFFF";
            final String LGRAY   = "D9D9D9"; // baris Istirahat / Sholat

            // ----- Style factory -----
            // Semua cell pakai border tipis, wrap, center
            Font fNormal = wb.createFont(); fNormal.setFontName("Arial"); fNormal.setFontHeightInPoints((short)9);
            Font fBold   = wb.createFont(); fBold.setFontName("Arial");   fBold.setFontHeightInPoints((short)9); fBold.setBold(true);
            Font fBoldWhite = wb.createFont(); fBoldWhite.setFontName("Arial"); fBoldWhite.setFontHeightInPoints((short)9); fBoldWhite.setBold(true); fBoldWhite.setColor(IndexedColors.WHITE.getIndex());

            Sheet sheet = wb.createSheet("Jadwal");
            sheet.setDefaultRowHeightInPoints(14);

            String[] namaHari = {"Senin", "Selasa", "Rabu", "Kamis", "Jumat"};

            // Nama kelas sesuai gambar: 7A-7H (8), 8A-8H (8), 9A-9G (7)
            // Nama kelas dinamis dari input UI
            int n7 = SchedulerUI.kelasT7;
            int n8 = SchedulerUI.kelasT8;
            int n9 = SchedulerUI.kelasT9;

            String[] kelas7 = new String[n7];
            for (int i = 0; i < n7; i++) kelas7[i] = "7" + (char)('A' + i);

            String[] kelas8 = new String[n8];
            for (int i = 0; i < n8; i++) kelas8[i] = "8" + (char)('A' + i);

            String[] kelas9 = new String[n9];
            for (int i = 0; i < n9; i++) kelas9[i] = "9" + (char)('A' + i);

            int totalKelas = n7 + n8 + n9;

            // Waktu jam ke- untuk tiap baris jadwal (index di jadwal[])
            // Kita pakai label waktu sesuai gambar
            String[] waktuLabel = {
                    "06.30 - 07.15", // SHO (Sholat Duha) -- bukan jam KBM
                    "07.15 - 07.55", // Jam 1
                    "07.55 - 08.35", // Jam 2
                    "08.35 - 09.15", // Jam 3
                    "09.15 - 09.55", // Jam 4
                    // ISTIRAHAT
                    "10.25 - 11.05", // Jam 5
                    "11.05 - 11.45", // Jam 6
                    // ISTIRAHAT / SHOLAT JUMAT
                    "12.20 - 13.00", // Jam 7
                    "13.00 - 13.40", // Jam 8
                    "13.40 - 14.20", // Jam 9
                    "14.20 - 15.00", // Jam 10
            };
            // "Jam ke" label (null = bukan jam KBM)
            String[] jamKeLabel = {"SHOLAT DHUHA","1","2","3","4",null,"5","6",null,"7","8","9","10"};

            // Mapping guru: hitung berapa jam per hari
            // Key = nomor guru (string angka), value = int[5] (hari 1-5)
            Map<String, int[]> guruJam = new HashMap<>();

            // Kumpulkan semua nomor guru
            for (String[] row : jadwal) {
                for (String cell : row) {
                    if (cell != null && !cell.isEmpty()) {
                        String num = cell.replaceAll("[^0-9]", "");
                        if (!num.isEmpty()) guruJam.putIfAbsent(num, new int[5]);
                    }
                }
            }
            // Hitung jam per hari
            for (int h = 0; h < hariRange.length; h++) {
                int start = hariRange[h][0];
                int end   = hariRange[h][1];
                for (int r = start; r <= end; r++) {
                    for (String cell : jadwal[r]) {
                        if (cell != null && !cell.isEmpty()) {
                            String num = cell.replaceAll("[^0-9]", "");
                            if (!num.isEmpty()) {
                                guruJam.computeIfAbsent(num, k -> new int[5]);
                                guruJam.get(num)[h]++;
                            }
                        }
                    }
                }
            }

            // Urutkan guru berdasarkan nomor
            List<String> guruList = new ArrayList<>(guruJam.keySet());
            guruList.sort(Comparator.comparingInt(Integer::parseInt));

            int COL_HARI  = 0;
            int COL_WAKTU = 1;
            int COL_JAM   = 2;
            int COL_KELAS_START = 3;
            int COL_KD    = COL_KELAS_START + totalKelas;
            int COL_NAMA  = COL_KD + 1;
            int COL_JML   = COL_NAMA + 1;
            int COL_HARI1 = COL_JML + 1; // hari 1..5 = COL_HARI1..COL_HARI1+4

            int rowIdx = 0;

            // Row 0: "HARI" | "WAKTU" | "JAM KE" | KELAS (merge) x3 | KD | NAMA GURU | JML | HARI (merge)
            Row hdr0 = sheet.createRow(rowIdx);
            hdr0.setHeightInPoints(16);

            setCell(wb, hdr0, COL_HARI,  "HARI",    fBold, WHITE,  HorizontalAlignment.CENTER, true);
            setCell(wb, hdr0, COL_WAKTU, "WAKTU",   fBold, WHITE,  HorizontalAlignment.CENTER, true);
            setCell(wb, hdr0, COL_JAM,   "JAM\nKE", fBold, WHITE,  HorizontalAlignment.CENTER, true);

            // Merge "KELAS" header atas tiap grup
            // Grup 7: col 3..10
            setCell(wb, hdr0, COL_KELAS_START,               "KELAS", fBold, GREEN,  HorizontalAlignment.CENTER, true);
            sheet.addMergedRegion(new CellRangeAddress(rowIdx, rowIdx, COL_KELAS_START, COL_KELAS_START + kelas7.length - 1));

            // Grup 8: col 11..18
            setCell(wb, hdr0, COL_KELAS_START + kelas7.length, "KELAS", fBold, YELLOW, HorizontalAlignment.CENTER, true);
            sheet.addMergedRegion(new CellRangeAddress(rowIdx, rowIdx, COL_KELAS_START + kelas7.length, COL_KELAS_START + kelas7.length + kelas8.length - 1));

            // Grup 9: col 19..25
            setCell(wb, hdr0, COL_KELAS_START + kelas7.length + kelas8.length, "KELAS", fBold, RED, HorizontalAlignment.CENTER, true);
            sheet.addMergedRegion(new CellRangeAddress(rowIdx, rowIdx, COL_KELAS_START + kelas7.length + kelas8.length, COL_KELAS_START + totalKelas - 1));

            setCell(wb, hdr0, COL_KD,   "KD",        fBold, LBLUE, HorizontalAlignment.CENTER, true);
            setCell(wb, hdr0, COL_NAMA, "NAMA\nGURU",fBold, LBLUE, HorizontalAlignment.CENTER, true);
            setCell(wb, hdr0, COL_JML,  "JML",       fBold, LBLUE, HorizontalAlignment.CENTER, true);

            setCell(wb, hdr0, COL_HARI1, "HARI", fBold, LBLUE, HorizontalAlignment.CENTER, true);
            sheet.addMergedRegion(new CellRangeAddress(rowIdx, rowIdx, COL_HARI1, COL_HARI1 + 4));

            rowIdx++;

            // Row 1: sub-header nama kelas + KD/NAMA/JML/1/2/3/4/5
            Row hdr1 = sheet.createRow(rowIdx);
            hdr1.setHeightInPoints(14);

            // HARI/WAKTU/JAMKE merge dua baris (0..1)
            sheet.addMergedRegion(new CellRangeAddress(0, 1, COL_HARI,  COL_HARI));
            sheet.addMergedRegion(new CellRangeAddress(0, 1, COL_WAKTU, COL_WAKTU));
            sheet.addMergedRegion(new CellRangeAddress(0, 1, COL_JAM,   COL_JAM));
            sheet.addMergedRegion(new CellRangeAddress(0, 1, COL_KD,    COL_KD));
            sheet.addMergedRegion(new CellRangeAddress(0, 1, COL_NAMA,  COL_NAMA));
            sheet.addMergedRegion(new CellRangeAddress(0, 1, COL_JML,   COL_JML));

            for (int i = 0; i < kelas7.length; i++)
                setCell(wb, hdr1, COL_KELAS_START + i, kelas7[i], fBold, GREEN, HorizontalAlignment.CENTER, true);
            for (int i = 0; i < kelas8.length; i++)
                setCell(wb, hdr1, COL_KELAS_START + kelas7.length + i, kelas8[i], fBold, YELLOW, HorizontalAlignment.CENTER, true);
            for (int i = 0; i < kelas9.length; i++)
                setCell(wb, hdr1, COL_KELAS_START + kelas7.length + kelas8.length + i, kelas9[i], fBold, RED, HorizontalAlignment.CENTER, true);

            for (int d = 1; d <= 5; d++)
                setCell(wb, hdr1, COL_HARI1 + d - 1, String.valueOf(d), fBold, LBLUE, HorizontalAlignment.CENTER, true);

            rowIdx++;

            // =====================================================================
            // ISI JADWAL per HARI
            // =====================================================================
            // Kita juga perlu tahu max baris per hari untuk merge kolom guru
            int guruRowStart = rowIdx; // baris mulai isi guru (kolom kanan)

            // Struktur baris tiap hari: Sholat Duha + 4 jam + Istirahat + 2 jam + Istirahat/Jumat + 4 jam = 12 KBM + 2/3 istirahat + 1 duha
            // Sesuai gambar: 13 baris per hari (termasuk header hari? tidak, header hari tidak ada di gambar utama)
            // Dari gambar: tiap hari punya baris Sholat Duha + Jam1..4 + Istirahat + Jam5..6 + Istirahat + Jam7..10 = 1+4+1+2+1+4 = 13 baris
            // Tapi di gambar Senin hanya ada 10 jam (baris Jam 1..10)

            // Hitung jumlah baris per hari (untuk merge HARI)
            // Tiap hari = 1(duha) + (end-start+1)(jam) + 1(istirahat setelah jam4) + 1(istirahat setelah jam6) = jam+3
            // Tapi baris HARI di-merge vertikal

            // Kita iterasi per hari
            for (int h = 0; h < hariRange.length; h++) {
                int start = hariRange[h][0];
                int end   = hariRange[h][1];
                int jamCount = end - start + 1;

                int hariRowStart = rowIdx;

                // Baris Sholat Duha
                Row duhaRow = sheet.createRow(rowIdx);
                duhaRow.setHeightInPoints(14);
                setCell(wb, duhaRow, COL_WAKTU, "06.30 - 07.15", fNormal, WHITE, HorizontalAlignment.CENTER, false);
                setCell(wb, duhaRow, COL_JAM,   "SHOLAT DHUHA",           fBold,   WHITE, HorizontalAlignment.CENTER, false);
                for (int k = 0; k < totalKelas; k++)
                    setCell(wb, duhaRow, COL_KELAS_START + k, "", fNormal, WHITE, HorizontalAlignment.CENTER, false);
                // Kolom kanan kosong (akan diisi guru nanti)
                rowIdx++;

                // Waktu jam-ke
                String[] waktuJam = {
                        "07.15 - 07.55","07.55 - 08.35","08.35 - 09.15","09.15 - 09.55",
                        "10.25 - 11.05","11.05 - 11.45",
                        "12.20 - 13.00","13.00 - 13.40","13.40 - 14.20","14.20 - 15.00"
                };
                // Istirahat muncul setelah jam 4 (index 3) dan setelah jam 6 (index 5)

                int jamKe = 0;
                for (int r = start; r <= end; r++) {
                    // Istirahat setelah jam ke-4
                    if (jamKe == 4) {
                        Row ist = sheet.createRow(rowIdx);
                        ist.setHeightInPoints(14);
                        setCell(wb, ist, COL_WAKTU, "09.55 - 10.25", fBold, LGRAY, HorizontalAlignment.CENTER, false);
                        setCell(wb, ist, COL_JAM,   "",              fBold, LGRAY, HorizontalAlignment.CENTER, false);
                        Cell istMerge = ist.createCell(COL_KELAS_START);
                        istMerge.setCellValue("ISTIRAHAT");
                        CellStyle istStyle = createStyle(wb, fBold, LGRAY, HorizontalAlignment.CENTER, true, true);
                        istMerge.setCellStyle(istStyle);
                        sheet.addMergedRegion(new CellRangeAddress(rowIdx, rowIdx, COL_KELAS_START, COL_KELAS_START + totalKelas - 1));
                        rowIdx++;
                    }
                    // Istirahat/Sholat Jumat setelah jam ke-6
                    if (jamKe == 6) {
                        String label = (h == 4) ? "SHOLAT JUMAT" : "SHOLAT ZUHUR DAN KULTUM";
                        Row ist2 = sheet.createRow(rowIdx);
                        ist2.setHeightInPoints(14);
                        setCell(wb, ist2, COL_WAKTU, "11.45 - 12.20", fBold, LGRAY, HorizontalAlignment.CENTER, false);
                        setCell(wb, ist2, COL_JAM,   "",               fBold, LGRAY, HorizontalAlignment.CENTER, false);
                        Cell ist2Merge = ist2.createCell(COL_KELAS_START);
                        ist2Merge.setCellValue(label);
                        ist2Merge.setCellStyle(createStyle(wb, fBold, LGRAY, HorizontalAlignment.CENTER, true, true));
                        sheet.addMergedRegion(new CellRangeAddress(rowIdx, rowIdx, COL_KELAS_START, COL_KELAS_START + totalKelas - 1));
                        rowIdx++;
                    }

                    Row row = sheet.createRow(rowIdx);
                    row.setHeightInPoints(14);
                    String wkt = (jamKe < waktuJam.length) ? waktuJam[jamKe] : "";
                    setCell(wb, row, COL_WAKTU, wkt,              fNormal, WHITE, HorizontalAlignment.CENTER, false);
                    setCell(wb, row, COL_JAM,   String.valueOf(jamKe + 1), fBold, WHITE, HorizontalAlignment.CENTER, false);
                    for (int k = 0; k < totalKelas; k++) {
                        String val = (r < jadwal.length && k < jadwal[r].length) ? jadwal[r][k] : "";
                        if (val == null) val = "";
                        setCell(wb, row, COL_KELAS_START + k, val, fNormal, WHITE, HorizontalAlignment.CENTER, false);
                    }
                    rowIdx++;
                    jamKe++;
                }

                int hariRowEnd = rowIdx - 1;

                // Merge kolom HARI vertikal
                Row firstHariRow = sheet.getRow(hariRowStart);
                setCell(wb, firstHariRow, COL_HARI, namaHari[h], fBold, WHITE, HorizontalAlignment.CENTER, true);
                if (hariRowEnd > hariRowStart)
                    sheet.addMergedRegion(new CellRangeAddress(hariRowStart, hariRowEnd, COL_HARI, COL_HARI));
            }

            // =====================================================================
            // KOLOM KANAN: KD | NAMA GURU | JML | HARI 1-5
            // Dimulai dari baris rowIdx=2 (setelah header 2 baris)
            // Tiap guru = 1 baris
            // =====================================================================
            int guruIdx = 0;
            for (String guru : guruList) {
                int dataRow = guruRowStart + guruIdx;
                Row row = sheet.getRow(dataRow);
                if (row == null) row = sheet.createRow(dataRow);

                int[] jams = guruJam.get(guru);
                int total = Arrays.stream(jams).sum();

                setCell(wb, row, COL_KD,   String.valueOf(guruIdx + 1), fNormal, WHITE, HorizontalAlignment.CENTER, false);
                String namaGuruDisplay = "";
                for (String[] data : kebutuhan) {
                    if (data[0].replaceAll("[^0-9]", "").equals(guru)) {
                        namaGuruDisplay = data[1];
                        break;
                    }
                }
                setCell(wb, row, COL_NAMA, namaGuruDisplay, fNormal, WHITE, HorizontalAlignment.LEFT, false);
                setCell(wb, row, COL_JML,  String.valueOf(total),        fNormal, WHITE, HorizontalAlignment.CENTER, false);

                // Max jam per hari untuk highlight
                int maxJam = Arrays.stream(jams).max().orElse(0);

                for (int d = 0; d < 5; d++) {
                    String bg = (jams[d] == maxJam && maxJam > 0 && jams[d] >= 8) ? ORANGE : WHITE;
                    setCell(wb, row, COL_HARI1 + d, String.valueOf(jams[d]), fNormal, bg, HorizontalAlignment.CENTER, false);
                }
                guruIdx++;
            }

            // =====================================================================
            // LEBAR KOLOM
            // =====================================================================
            sheet.setColumnWidth(COL_HARI,  12 * 256);
            sheet.setColumnWidth(COL_WAKTU, 16 * 256);
            sheet.setColumnWidth(COL_JAM,    5 * 256);
            for (int k = 0; k < totalKelas; k++)
                sheet.setColumnWidth(COL_KELAS_START + k, 5 * 256);
            sheet.setColumnWidth(COL_KD,    4  * 256);
            sheet.setColumnWidth(COL_NAMA,  16 * 256);
            sheet.setColumnWidth(COL_JML,   5  * 256);
            for (int d = 0; d < 5; d++)
                sheet.setColumnWidth(COL_HARI1 + d, 4 * 256);

            // =====================================================================
            // FREEZE panes: bekukan 2 baris header + 3 kolom kiri
            // =====================================================================
            sheet.createFreezePane(3, 2);

            exportRekapMapel((XSSFWorkbook) wb, jadwal, kebutuhan);
            exportSheetGuru((XSSFWorkbook) wb, jadwal, hariRange, kebutuhan);
            exportSheetKelas((XSSFWorkbook) wb, jadwal, hariRange, kebutuhan);

            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                wb.write(fos);
            }
        }
        if (SchedulerUI.stopRequested) {
            SchedulerUI.markDone(false, "» Dihentikan manual — solusi terakhir disimpan.");
        } else {
            SchedulerUI.markDone(true, null);
        }
    }

    // =====================================================================
    // HELPER: set cell value + style
    // =====================================================================
    private static void setCell(Workbook wb, Row row, int col, String value,
                                Font font, String bgHex,
                                HorizontalAlignment ha, boolean wrap) {
        Cell cell = row.getCell(col);
        if (cell == null) cell = row.createCell(col);
        cell.setCellValue(value);
        cell.setCellStyle(createStyle(wb, font, bgHex, ha, wrap, true));
    }

    static void exportSheetGuru(XSSFWorkbook wb, String[][] jadwal, int[][] hariRange,
                                List<String[]> kebutuhan) throws Exception {

        final String WHITE = "FFFFFF";
        final String LGRAY = "F2F2F2";
        final String CYAN  = "B2EBF2";

        final String[] WAKTU_BARIS = {
                "06.30 - 07.15","07.15 - 07.55","07.55 - 08.35","08.35 - 09.15","09.15 - 09.55",
                "09.55 - 10.25","10.25 - 11.05","11.05 - 11.45","11.45 - 12.20",
                "12.20 - 13.00","13.00 - 13.40","13.40 - 14.20","14.20 - 15.00"
        };
        final String[] JAM_KE     = {null,"1","2","3","4",null,"5","6",null,"7","8","9","10"};
        final String[] LBL_KHUSUS = {"Sholat Dhuha",null,null,null,null,"Istirahat",null,null,"Sholat Zuhur",null,null,null,null};
        final String[] NAMA_HARI  = {"SENIN","SELASA","RABU","KAMIS","JUMAT"};
        final int barisPer = WAKTU_BARIS.length; // 13

        String[] allKelas = buatDaftarKelas();

        // Kumpulkan guru unik
        Map<String, String[]> guruMap = new LinkedHashMap<>();
        for (String[] data : kebutuhan) {
            String num = data[0].replaceAll("[^0-9]", "");
            if (!guruMap.containsKey(num))
                guruMap.put(num, new String[]{data[1], data[2]});
        }
        List<String> guruList = new ArrayList<>(guruMap.keySet());
        guruList.sort(Comparator.comparingInt(Integer::parseInt));

        // Map guruNum+kelasIdx -> mapel (untuk lookup mata pelajaran per kelas)
        Map<String, String> guruKelasMapel = new HashMap<>();
        for (String[] data : kebutuhan) {
            String num = data[0].replaceAll("[^0-9]","");
            try {
                int kIdx = Integer.parseInt(data[3]);
                guruKelasMapel.put(num+"|"+kIdx, data[2]);
            } catch (Exception ignored) {}
        }

        // Hapus sheet lama jika ada
        int ex = wb.getSheetIndex("Jadwal Guru");
        if (ex >= 0) wb.removeSheetAt(ex);
        ex = wb.getSheetIndex("DataGuru");
        if (ex >= 0) wb.removeSheetAt(ex);
        ex = wb.getSheetIndex("VBACode");
        if (ex >= 0) wb.removeSheetAt(ex);

        // Buat satu sheet untuk setiap guru
        for (String guruNum : guruList) {
            String[] info = guruMap.get(guruNum);
            String namaGuru = info[0];

            // Hitung total jam mengajar guru ini
            int totalJam = 0;
            for (int h = 0; h < hariRange.length; h++) {
                int start = hariRange[h][0], end = hariRange[h][1];
                for (int r = start; r <= end; r++) {
                    for (int k = 0; k < allKelas.length; k++) {
                        if (k < jadwal[r].length) {
                            String isi = jadwal[r][k];
                            if (isi != null && !isi.isEmpty() && isi.replaceAll("[^0-9]","").equals(guruNum)) {
                                totalJam++;
                                break;
                            }
                        }
                    }
                }
            }

            String sheetName = "Guru " + guruNum + " - " + namaGuru;
            int si = wb.getSheetIndex(sheetName);
            if (si >= 0) wb.removeSheetAt(si);
            XSSFSheet sheet = wb.createSheet(sheetName);
            sheet.setDefaultRowHeightInPoints(14);

            XSSFFont fN = guruFont(wb, 10, false);
            XSSFFont fB = guruFont(wb, 10, true);

            int[] cw = {8,5,16,10,30};
            for (int i = 0; i < cw.length; i++) sheet.setColumnWidth(i, cw[i]*256);

            int rowIdx = 0;

            // Header info guru (3 baris)
            Row r1 = sheet.createRow(rowIdx++); r1.setHeightInPoints(20);
            guruCell(wb, r1, 0, "KODE GURU", fB, WHITE, HorizontalAlignment.LEFT, false, false);
            guruCell(wb, r1, 1, ":",          fB, WHITE, HorizontalAlignment.CENTER, false, false);
            guruCell(wb, r1, 2, guruNum,      fB, WHITE, HorizontalAlignment.LEFT, false, false);

            Row r2 = sheet.createRow(rowIdx++); r2.setHeightInPoints(20);
            guruCell(wb, r2, 0, "NAMA GURU", fB, WHITE, HorizontalAlignment.LEFT, false, false);
            guruCell(wb, r2, 1, ":",          fB, WHITE, HorizontalAlignment.CENTER, false, false);
            guruCell(wb, r2, 2, namaGuru,     fB, WHITE, HorizontalAlignment.LEFT, false, false);

            Row r3 = sheet.createRow(rowIdx++); r3.setHeightInPoints(20);
            guruCell(wb, r3, 0, "JUMLAH MENGAJAR", fB, WHITE, HorizontalAlignment.LEFT, false, false);
            guruCell(wb, r3, 1, ":",               fB, WHITE, HorizontalAlignment.CENTER, false, false);
            guruCell(wb, r3, 2, totalJam + " JAM", fB, WHITE, HorizontalAlignment.LEFT, false, false);

            rowIdx++; // spasi

            // Header tabel
            Row rH = sheet.createRow(rowIdx++); rH.setHeightInPoints(16);
            guruCell(wb, rH, 0, "HARI",          fB, LGRAY, HorizontalAlignment.CENTER, true, true);
            guruCell(wb, rH, 1, "JAM",           fB, LGRAY, HorizontalAlignment.CENTER, true, true);
            guruCell(wb, rH, 2, "WAKTU",         fB, LGRAY, HorizontalAlignment.CENTER, true, true);
            guruCell(wb, rH, 3, "KLS",           fB, LGRAY, HorizontalAlignment.CENTER, true, true);
            guruCell(wb, rH, 4, "MATA PELAJARAN",fB, LGRAY, HorizontalAlignment.CENTER, true, true);

            // Tabel jadwal per hari
            for (int h = 0; h < hariRange.length; h++) {
                int start    = hariRange[h][0];
                int end      = hariRange[h][1];
                int jamAktual = end - start + 1;

                int startRow = rowIdx;
                int jamKe = 0;
                int bRow  = 0;

                for (int b = 0; b < barisPer; b++) {
                    if (JAM_KE[b] != null) {
                        if (jamKe >= jamAktual) break;
                        jamKe++;
                    } else {
                        boolean adaJamBerikut = false;
                        for (int nb = b+1; nb < barisPer; nb++) {
                            if (JAM_KE[nb] != null && jamKe < jamAktual) { adaJamBerikut = true; break; }
                        }
                        if (!adaJamBerikut) break;
                    }

                    Row row = sheet.createRow(rowIdx++);
                    row.setHeightInPoints(14);

                    if (LBL_KHUSUS[b] != null) {
                        guruCell(wb, row, 0, "",               fN, CYAN, HorizontalAlignment.CENTER, true, true);
                        guruCell(wb, row, 1, "",               fN, CYAN, HorizontalAlignment.CENTER, true, true);
                        guruCell(wb, row, 2, WAKTU_BARIS[b],   fN, CYAN, HorizontalAlignment.CENTER, true, true);
                        guruCell(wb, row, 3, LBL_KHUSUS[b],    fB, CYAN, HorizontalAlignment.CENTER, true, true);
                        sheet.addMergedRegion(new CellRangeAddress(rowIdx-1, rowIdx-1, 3, 4));
                        guruCell(wb, row, 4, "",               fN, CYAN, HorizontalAlignment.CENTER, true, true);
                    } else {
                        int slot = start + jamKe - 1;
                        String kelas = "";
                        String mapel = "";
                        for (int k = 0; k < allKelas.length; k++) {
                            if (slot < jadwal.length && k < jadwal[slot].length) {
                                String isi = jadwal[slot][k];
                                if (isi != null && !isi.isEmpty() && isi.replaceAll("[^0-9]","").equals(guruNum)) {
                                    kelas = allKelas[k];
                                    mapel = guruKelasMapel.getOrDefault(guruNum + "|" + k, "");
                                    break;
                                }
                            }
                        }
                        guruCell(wb, row, 0, "",        fN, WHITE, HorizontalAlignment.CENTER, true, true);
                        guruCell(wb, row, 1, JAM_KE[b], fN, WHITE, HorizontalAlignment.CENTER, true, true);
                        guruCell(wb, row, 2, WAKTU_BARIS[b], fN, WHITE, HorizontalAlignment.CENTER, true, true);
                        guruCell(wb, row, 3, kelas,    fN, WHITE, HorizontalAlignment.CENTER, true, true);
                        guruCell(wb, row, 4, mapel,    fN, WHITE, HorizontalAlignment.LEFT, true, true);
                    }
                    bRow++;
                }

                // Merge kolom HARI vertikal
                if (startRow < rowIdx) {
                    sheet.addMergedRegion(new CellRangeAddress(startRow, rowIdx-1, 0, 0));
                    sheet.getRow(startRow).getCell(0).setCellValue(NAMA_HARI[h]);
                    sheet.getRow(startRow).getCell(0).setCellStyle(guruStyle(wb, fB, LGRAY, HorizontalAlignment.CENTER, true, true));
                }

                // Spasi antar hari
                sheet.createRow(rowIdx++).setHeightInPoints(6);
            }
        }
    }
    // Helper: hitung berapa baris yang dipakai untuk sejumlah jam aktual
    private static int getBarisPakai(int jamAktual, String[] JAM_KE, int barisPer) {
        int jamKe  = 0;
        int baris  = 0;
        for (int b = 0; b < barisPer; b++) {
            if (JAM_KE[b] != null) {
                if (jamKe >= jamAktual) break; // jam sudah cukup, stop
                jamKe++;
            } else {
                // Istirahat/sholat: hitung hanya jika masih ada jam setelahnya
                boolean adaJamBerikut = false;
                for (int nb = b+1; nb < barisPer; nb++) {
                    if (JAM_KE[nb] != null && jamKe < jamAktual) {
                        adaJamBerikut = true;
                        break;
                    }
                }
                if (!adaJamBerikut) break;
            }
            baris++;
        }
        return baris;
    }

    // =======================================================================
// SHEET JADWAL KELAS
// =======================================================================
    static void exportSheetKelas(XSSFWorkbook wb, String[][] jadwal, int[][] hariRange,
                                 List<String[]> kebutuhan) throws Exception {

        final String WHITE = "FFFFFF";
        final String LGRAY = "F2F2F2";
        final String CYAN  = "B2EBF2";

        final String[] WAKTU_BARIS = {
                "06.30 - 07.15","07.15 - 07.55","07.55 - 08.35","08.35 - 09.15","09.15 - 09.55",
                "09.55 - 10.25","10.25 - 11.05","11.05 - 11.45","11.45 - 12.20",
                "12.20 - 13.00","13.00 - 13.40","13.40 - 14.20","14.20 - 15.00"
        };
        final String[] JAM_KE     = {null,"1","2","3","4",null,"5","6",null,"7","8","9","10"};
        final String[] LBL_KHUSUS = {"Sholat Dhuha",null,null,null,null,"Istirahat",null,null,"Sholat Zuhur",null,null,null,null};
        final String[] NAMA_HARI  = {"SENIN","SELASA","RABU","KAMIS","JUMAT"};
        final int barisPer = WAKTU_BARIS.length;

        String[] allKelas = buatDaftarKelas();

        // Map guruNum -> namaGuru dan guruNum+kelasIdx -> mapel
        Map<String, String> guruNama = new HashMap<>();
        Map<String, String> guruKelasMapel = new HashMap<>();
        for (String[] data : kebutuhan) {
            String num = data[0].replaceAll("[^0-9]","");
            guruNama.putIfAbsent(num, data[1]);
            try {
                int kIdx = Integer.parseInt(data[3]);
                guruKelasMapel.put(num+"|"+kIdx, data[2]);
            } catch (Exception ignored) {}
        }

        // Hapus sheet lama jika ada
        int ex = wb.getSheetIndex("Jadwal Kelas");
        if (ex >= 0) wb.removeSheetAt(ex);
        ex = wb.getSheetIndex("DataKelas");
        if (ex >= 0) wb.removeSheetAt(ex);
        ex = wb.getSheetIndex("VBAKelas");
        if (ex >= 0) wb.removeSheetAt(ex);

        // Buat satu sheet untuk setiap kelas
        for (int kIdx = 0; kIdx < allKelas.length; kIdx++) {
            String namaKelas = allKelas[kIdx];

            String sheetName = "Kelas " + namaKelas;
            int si = wb.getSheetIndex(sheetName);
            if (si >= 0) wb.removeSheetAt(si);
            XSSFSheet sheet = wb.createSheet(sheetName);
            sheet.setDefaultRowHeightInPoints(14);

            XSSFFont fN = guruFont(wb, 10, false);
            XSSFFont fB = guruFont(wb, 10, true);

            int[] cw = {8,5,16,25,25};
            for (int i = 0; i < cw.length; i++) sheet.setColumnWidth(i, cw[i]*256);

            int rowIdx = 0;

            // Info kelas
            Row rK = sheet.createRow(rowIdx++); rK.setHeightInPoints(20);
            guruCell(wb, rK, 0, "KELAS", fB, WHITE, HorizontalAlignment.LEFT, false, false);
            guruCell(wb, rK, 1, ":",      fB, WHITE, HorizontalAlignment.CENTER, false, false);
            guruCell(wb, rK, 2, namaKelas, fB, WHITE, HorizontalAlignment.LEFT, false, false);

            rowIdx++; // spasi

            // Header tabel
            Row rH = sheet.createRow(rowIdx++); rH.setHeightInPoints(16);
            guruCell(wb, rH, 0, "HARI",      fB, LGRAY, HorizontalAlignment.CENTER, true, true);
            guruCell(wb, rH, 1, "JAM",       fB, LGRAY, HorizontalAlignment.CENTER, true, true);
            guruCell(wb, rH, 2, "WAKTU",     fB, LGRAY, HorizontalAlignment.CENTER, true, true);
            guruCell(wb, rH, 3, "MAPEL",     fB, LGRAY, HorizontalAlignment.CENTER, true, true);
            guruCell(wb, rH, 4, "NAMA GURU", fB, LGRAY, HorizontalAlignment.CENTER, true, true);

            // Tabel jadwal per hari
            for (int h = 0; h < hariRange.length; h++) {
                int start    = hariRange[h][0];
                int end      = hariRange[h][1];
                int jamAktual = end - start + 1;

                int startRow = rowIdx;
                int jamKe = 0;
                int bRow  = 0;

                for (int b = 0; b < barisPer; b++) {
                    if (JAM_KE[b] != null) {
                        if (jamKe >= jamAktual) break;
                        jamKe++;
                    } else {
                        boolean adaJamBerikut = false;
                        for (int nb = b+1; nb < barisPer; nb++) {
                            if (JAM_KE[nb] != null && jamKe < jamAktual) { adaJamBerikut = true; break; }
                        }
                        if (!adaJamBerikut) break;
                    }

                    Row row = sheet.createRow(rowIdx++);
                    row.setHeightInPoints(14);

                    if (LBL_KHUSUS[b] != null) {
                        guruCell(wb, row, 0, "",               fN, CYAN, HorizontalAlignment.CENTER, true, true);
                        guruCell(wb, row, 1, "",               fN, CYAN, HorizontalAlignment.CENTER, true, true);
                        guruCell(wb, row, 2, WAKTU_BARIS[b],   fN, CYAN, HorizontalAlignment.CENTER, true, true);
                        guruCell(wb, row, 3, LBL_KHUSUS[b],    fB, CYAN, HorizontalAlignment.CENTER, true, true);
                        sheet.addMergedRegion(new CellRangeAddress(rowIdx-1, rowIdx-1, 3, 4));
                        guruCell(wb, row, 4, "",               fN, CYAN, HorizontalAlignment.CENTER, true, true);
                    } else {
                        int slot = start + jamKe - 1;
                        String guruId = "";
                        String mapel = "";
                        String guruName = "";
                        if (slot < jadwal.length && kIdx < jadwal[slot].length) {
                            String isi = jadwal[slot][kIdx];
                            if (isi != null && !isi.isEmpty()) {
                                guruId = isi.replaceAll("[^0-9]","");
                                guruName = guruNama.getOrDefault(guruId, "");
                                mapel = guruKelasMapel.getOrDefault(guruId+"|"+kIdx, "");
                            }
                        }
                        guruCell(wb, row, 0, "",        fN, WHITE, HorizontalAlignment.CENTER, true, true);
                        guruCell(wb, row, 1, JAM_KE[b], fN, WHITE, HorizontalAlignment.CENTER, true, true);
                        guruCell(wb, row, 2, WAKTU_BARIS[b], fN, WHITE, HorizontalAlignment.CENTER, true, true);
                        guruCell(wb, row, 3, mapel,    fN, WHITE, HorizontalAlignment.CENTER, true, true);
                        guruCell(wb, row, 4, guruName, fN, WHITE, HorizontalAlignment.LEFT, true, true);
                    }
                    bRow++;
                }

                // Merge kolom HARI vertikal
                if (startRow < rowIdx) {
                    sheet.addMergedRegion(new CellRangeAddress(startRow, rowIdx-1, 0, 0));
                    sheet.getRow(startRow).getCell(0).setCellValue(NAMA_HARI[h]);
                    sheet.getRow(startRow).getCell(0).setCellStyle(guruStyle(wb, fB, LGRAY, HorizontalAlignment.CENTER, true, true));
                }

                // Spasi antar hari
                sheet.createRow(rowIdx++).setHeightInPoints(6);
            }
        }
    }

    // =======================================================================
    // REKAP MAPEL — sheet ringkasan jam mengajar per mapel per kelas
    // =======================================================================
    static void exportRekapMapel(XSSFWorkbook wb, String[][] jadwal, List<String[]> kebutuhan) throws Exception {
        final String WHITE = "FFFFFF";
        final String LGRAY = "F2F2F2";
        final String BLUE  = "2563EB";

        // Build: guruNum|kelasIdx -> subject
        Map<String, String> guruKelasMapel = new HashMap<>();
        for (String[] d : kebutuhan) {
            String num = d[0].replaceAll("[^0-9]", "");
            guruKelasMapel.put(num + "|" + d[3], d[2]);
        }

        // Also collect guruNum -> namaGuru
        Map<String, String> guruNama = new HashMap<>();
        for (String[] d : kebutuhan) {
            String num = d[0].replaceAll("[^0-9]", "");
            guruNama.putIfAbsent(num, d[1]);
        }

        // Count unique teachers: subject -> Set<String>[3] (teacher IDs per grade 7,8,9)
        Map<String, Set<String>[]> teacherMap = new LinkedHashMap<>();
        int totalKelas = jadwal[0].length;
        int k7 = SchedulerUI.kelasT7;
        int k8 = SchedulerUI.kelasT8;

        for (int r = 0; r < jadwal.length; r++) {
            for (int k = 0; k < totalKelas; k++) {
                String cell = jadwal[r][k];
                if (cell == null || cell.isEmpty()) continue;
                String num = cell.replaceAll("[^0-9]", "");
                String subject = guruKelasMapel.get(num + "|" + k);
                if (subject == null) continue;
                int gradeIdx;
                if (k < k7) gradeIdx = 0;
                else if (k < k7 + k8) gradeIdx = 1;
                else gradeIdx = 2;
                teacherMap.computeIfAbsent(subject, s -> {
                    @SuppressWarnings("unchecked")
                    Set<String>[] sets = new Set[]{new HashSet<>(), new HashSet<>(), new HashSet<>()};
                    return sets;
                })[gradeIdx].add(num);
            }
        }

        XSSFSheet sheet = wb.createSheet("Rekap Mapel");
        sheet.setDefaultRowHeightInPoints(14);

        XSSFFont fN = guruFont(wb, 10, false);
        XSSFFont fB = guruFont(wb, 10, true);
        XSSFFont fBW = guruFont(wb, 10, true);
        fBW.setColor(IndexedColors.WHITE.getIndex());

        int[] cw = {4, 28, 20, 20, 20};
        String[] cols = {"NO", "MAPEL", "KELAS 7", "KELAS 8", "KELAS 9"};
        for (int i = 0; i < cw.length; i++) sheet.setColumnWidth(i, cw[i] * 256);

        int rowIdx = 0;

        // Title bar
        Row rTitle = sheet.createRow(rowIdx++); rTitle.setHeightInPoints(22);
        guruCell(wb, rTitle, 0, "REKAP JUMLAH GURU PER MAPEL", fBW, BLUE, HorizontalAlignment.LEFT, false, false);
        for (int i = 1; i < cols.length; i++) {
            guruCell(wb, rTitle, i, "", fN, BLUE, HorizontalAlignment.CENTER, false, false);
        }
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 4));

        rowIdx++;

        // Table header
        Row rH = sheet.createRow(rowIdx++);
        for (int i = 0; i < cols.length; i++) {
            guruCell(wb, rH, i, cols[i], fB, LGRAY, HorizontalAlignment.CENTER, true, true);
        }

        // Integer format for numeric columns
        XSSFCellStyle intStyle = guruStyle(wb, fN, WHITE, HorizontalAlignment.CENTER, true, true);
        intStyle.setDataFormat(wb.createDataFormat().getFormat("0"));

        // Data rows
        int no = 1;
        for (Map.Entry<String, Set<String>[]> entry : teacherMap.entrySet()) {
            Set<String>[] sets = entry.getValue();
            int total = sets[0].size() + sets[1].size() + sets[2].size();
            if (total == 0) continue;
            Row row = sheet.createRow(rowIdx++);
            guruCell(wb, row, 0, String.valueOf(no++), fN, WHITE, HorizontalAlignment.CENTER, true, true);
            guruCell(wb, row, 1, entry.getKey(),       fN, WHITE, HorizontalAlignment.LEFT, true, true);
            row.createCell(2).setCellValue(sets[0].size()); row.getCell(2).setCellStyle(intStyle);
            row.createCell(3).setCellValue(sets[1].size()); row.getCell(3).setCellStyle(intStyle);
            row.createCell(4).setCellValue(sets[2].size()); row.getCell(4).setCellStyle(intStyle);
        }
    }

    // =======================================================================
    // PREVIEW DATA (for SchedulerUI [Cetak] panel)
    // =======================================================================
    static Object[] generateGuruPreview(String guruNum, String[][] jadwal, int[][] hariRange,
                                        List<String[]> kebutuhan) {
        final String[] WAKTU_BARIS = {
            "06.30 - 07.15","07.15 - 07.55","07.55 - 08.35","08.35 - 09.15","09.15 - 09.55",
            "09.55 - 10.25","10.25 - 11.05","11.05 - 11.45","11.45 - 12.20",
            "12.20 - 13.00","13.00 - 13.40","13.40 - 14.20","14.20 - 15.00"
        };
        final String[] JAM_KE     = {null,"1","2","3","4",null,"5","6",null,"7","8","9","10"};
        final String[] LBL_KHUSUS = {"Sholat Dhuha",null,null,null,null,"Istirahat",null,null,"Sholat Zuhur",null,null,null,null};
        final String[] NAMA_HARI  = {"SENIN","SELASA","RABU","KAMIS","JUMAT"};
        final int barisPer = WAKTU_BARIS.length;

        String[] allKelas = buatDaftarKelas();

        Map<String, String> guruKelasMapel = new HashMap<>();
        for (String[] data : kebutuhan) {
            String num = data[0].replaceAll("[^0-9]","");
            try {
                int kIdx = Integer.parseInt(data[3]);
                guruKelasMapel.put(num+"|"+kIdx, data[2]);
            } catch (Exception ignored) {}
        }

        java.util.List<String[]> rows = new java.util.ArrayList<>();
        rows.add(new String[]{"HARI","JAM","WAKTU","KLS","MATA PELAJARAN"});

        for (int h = 0; h < hariRange.length; h++) {
            int start = hariRange[h][0], end = hariRange[h][1];
            int jamAktual = end - start + 1;
            int jamKe = 0;

            for (int b = 0; b < barisPer; b++) {
                if (JAM_KE[b] != null) {
                    if (jamKe >= jamAktual) break;
                    jamKe++;
                } else {
                    boolean adaJamBerikut = false;
                    for (int nb = b+1; nb < barisPer; nb++) {
                        if (JAM_KE[nb] != null && jamKe < jamAktual) { adaJamBerikut = true; break; }
                    }
                    if (!adaJamBerikut) break;
                }

                String hariLabel = NAMA_HARI[h];
                String jamLabel  = JAM_KE[b] != null ? JAM_KE[b] : "";
                String waktu     = WAKTU_BARIS[b];
                String kelas     = "";
                String mapel     = "";

                if (LBL_KHUSUS[b] != null) {
                    kelas = LBL_KHUSUS[b];
                } else {
                    int slot = start + jamKe - 1;
                    for (int k = 0; k < allKelas.length; k++) {
                        if (slot < jadwal.length && k < jadwal[slot].length) {
                            String isi = jadwal[slot][k];
                            if (isi != null && !isi.isEmpty() && isi.replaceAll("[^0-9]","").equals(guruNum)) {
                                kelas = allKelas[k];
                                mapel = guruKelasMapel.getOrDefault(guruNum + "|" + k, "");
                                break;
                            }
                        }
                    }
                }
                rows.add(new String[]{hariLabel, jamLabel, waktu, kelas, mapel});
            }
        }
        return new Object[]{rows.toArray(new String[0][]), new String[]{"HARI","JAM","WAKTU","KLS","MATA PELAJARAN"}};
    }

    static Object[] generateKelasPreview(int kelasIdx, String[][] jadwal, int[][] hariRange,
                                         List<String[]> kebutuhan) {
        final String[] WAKTU_BARIS = {
            "06.30 - 07.15","07.15 - 07.55","07.55 - 08.35","08.35 - 09.15","09.15 - 09.55",
            "09.55 - 10.25","10.25 - 11.05","11.05 - 11.45","11.45 - 12.20",
            "12.20 - 13.00","13.00 - 13.40","13.40 - 14.20","14.20 - 15.00"
        };
        final String[] JAM_KE     = {null,"1","2","3","4",null,"5","6",null,"7","8","9","10"};
        final String[] LBL_KHUSUS = {"Sholat Dhuha",null,null,null,null,"Istirahat",null,null,"Sholat Zuhur",null,null,null,null};
        final String[] NAMA_HARI  = {"SENIN","SELASA","RABU","KAMIS","JUMAT"};
        final int barisPer = WAKTU_BARIS.length;

        Map<String, String> guruNama = new HashMap<>();
        Map<String, String> guruKelasMapel = new HashMap<>();
        for (String[] data : kebutuhan) {
            String num = data[0].replaceAll("[^0-9]","");
            guruNama.putIfAbsent(num, data[1]);
            try {
                int kIdx = Integer.parseInt(data[3]);
                guruKelasMapel.put(num+"|"+kIdx, data[2]);
            } catch (Exception ignored) {}
        }

        java.util.List<String[]> rows = new java.util.ArrayList<>();
        rows.add(new String[]{"HARI","JAM","WAKTU","MAPEL","NAMA GURU"});

        for (int h = 0; h < hariRange.length; h++) {
            int start = hariRange[h][0], end = hariRange[h][1];
            int jamAktual = end - start + 1;
            int jamKe = 0;

            for (int b = 0; b < barisPer; b++) {
                if (JAM_KE[b] != null) {
                    if (jamKe >= jamAktual) break;
                    jamKe++;
                } else {
                    boolean adaJamBerikut = false;
                    for (int nb = b+1; nb < barisPer; nb++) {
                        if (JAM_KE[nb] != null && jamKe < jamAktual) { adaJamBerikut = true; break; }
                    }
                    if (!adaJamBerikut) break;
                }

                String hariLabel = NAMA_HARI[h];
                String jamLabel  = JAM_KE[b] != null ? JAM_KE[b] : "";
                String waktu     = WAKTU_BARIS[b];
                String mapel     = "";
                String guruName  = "";

                if (LBL_KHUSUS[b] != null) {
                    mapel    = LBL_KHUSUS[b];
                } else {
                    int slot = start + jamKe - 1;
                    if (slot < jadwal.length && kelasIdx < jadwal[slot].length) {
                        String isi = jadwal[slot][kelasIdx];
                        if (isi != null && !isi.isEmpty()) {
                            String guruId = isi.replaceAll("[^0-9]","");
                            guruName = guruNama.getOrDefault(guruId, "");
                            mapel = guruKelasMapel.getOrDefault(guruId+"|"+kelasIdx, "");
                        }
                    }
                }
                rows.add(new String[]{hariLabel, jamLabel, waktu, mapel, guruName});
            }
        }
        return new Object[]{rows.toArray(new String[0][]), new String[]{"HARI","JAM","WAKTU","MAPEL","NAMA GURU"}};
    }

    // =======================================================================
    // SINGLE EXPORT (for SchedulerUI [Excel] button)
    // =======================================================================
    static void exportSingleGuruToFile(String guruNum, String[][] jadwal, int[][] hariRange,
                                       List<String[]> kebutuhan, String outputPath) throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            final String WHITE = "FFFFFF";
            final String LGRAY = "F2F2F2";
            final String CYAN  = "B2EBF2";

            final String[] WAKTU_BARIS = {
                "06.30 - 07.15","07.15 - 07.55","07.55 - 08.35","08.35 - 09.15","09.15 - 09.55",
                "09.55 - 10.25","10.25 - 11.05","11.05 - 11.45","11.45 - 12.20",
                "12.20 - 13.00","13.00 - 13.40","13.40 - 14.20","14.20 - 15.00"
            };
            final String[] JAM_KE     = {null,"1","2","3","4",null,"5","6",null,"7","8","9","10"};
            final String[] LBL_KHUSUS = {"Sholat Dhuha",null,null,null,null,"Istirahat",null,null,"Sholat Zuhur",null,null,null,null};
            final String[] NAMA_HARI  = {"SENIN","SELASA","RABU","KAMIS","JUMAT"};
            final int barisPer = WAKTU_BARIS.length;

            String[] allKelas = buatDaftarKelas();
            String namaGuru = "";
            Map<String, String> guruKelasMapel = new HashMap<>();
            for (String[] data : kebutuhan) {
                String num = data[0].replaceAll("[^0-9]","");
                if (num.equals(guruNum) && namaGuru.isEmpty()) namaGuru = data[1];
                try {
                    int kIdx = Integer.parseInt(data[3]);
                    guruKelasMapel.put(num+"|"+kIdx, data[2]);
                } catch (Exception ignored) {}
            }

            int totalJam = 0;
            for (int h = 0; h < hariRange.length; h++) {
                int start = hariRange[h][0], end = hariRange[h][1];
                for (int r = start; r <= end; r++) {
                    for (int k = 0; k < allKelas.length; k++) {
                        if (k < jadwal[r].length) {
                            String isi = jadwal[r][k];
                            if (isi != null && !isi.isEmpty() && isi.replaceAll("[^0-9]","").equals(guruNum)) {
                                totalJam++;
                                break;
                            }
                        }
                    }
                }
            }

            String sheetName = "Guru " + guruNum + " - " + namaGuru;
            XSSFSheet sheet = wb.createSheet(sheetName);
            sheet.setDefaultRowHeightInPoints(14);

            XSSFFont fN = guruFont(wb, 10, false);
            XSSFFont fB = guruFont(wb, 10, true);
            int[] cw = {8,5,16,10,30};
            for (int i = 0; i < cw.length; i++) sheet.setColumnWidth(i, cw[i]*256);

            int rowIdx = 0;
            Row r1 = sheet.createRow(rowIdx++); r1.setHeightInPoints(20);
            guruCell(wb, r1, 0, "KODE GURU", fB, WHITE, HorizontalAlignment.LEFT, false, false);
            guruCell(wb, r1, 1, ":",          fB, WHITE, HorizontalAlignment.CENTER, false, false);
            guruCell(wb, r1, 2, guruNum,      fB, WHITE, HorizontalAlignment.LEFT, false, false);

            Row r2 = sheet.createRow(rowIdx++); r2.setHeightInPoints(20);
            guruCell(wb, r2, 0, "NAMA GURU", fB, WHITE, HorizontalAlignment.LEFT, false, false);
            guruCell(wb, r2, 1, ":",          fB, WHITE, HorizontalAlignment.CENTER, false, false);
            guruCell(wb, r2, 2, namaGuru,     fB, WHITE, HorizontalAlignment.LEFT, false, false);

            Row r3 = sheet.createRow(rowIdx++); r3.setHeightInPoints(20);
            guruCell(wb, r3, 0, "JUMLAH MENGAJAR", fB, WHITE, HorizontalAlignment.LEFT, false, false);
            guruCell(wb, r3, 1, ":",               fB, WHITE, HorizontalAlignment.CENTER, false, false);
            guruCell(wb, r3, 2, totalJam + " JAM", fB, WHITE, HorizontalAlignment.LEFT, false, false);

            rowIdx++;
            Row rH = sheet.createRow(rowIdx++); rH.setHeightInPoints(16);
            guruCell(wb, rH, 0, "HARI",          fB, LGRAY, HorizontalAlignment.CENTER, true, true);
            guruCell(wb, rH, 1, "JAM",           fB, LGRAY, HorizontalAlignment.CENTER, true, true);
            guruCell(wb, rH, 2, "WAKTU",         fB, LGRAY, HorizontalAlignment.CENTER, true, true);
            guruCell(wb, rH, 3, "KLS",           fB, LGRAY, HorizontalAlignment.CENTER, true, true);
            guruCell(wb, rH, 4, "MATA PELAJARAN",fB, LGRAY, HorizontalAlignment.CENTER, true, true);

            for (int h = 0; h < hariRange.length; h++) {
                int start = hariRange[h][0], end = hariRange[h][1];
                int jamAktual = end - start + 1;
                int startRow = rowIdx;
                int jamKe = 0;

                for (int b = 0; b < barisPer; b++) {
                    if (JAM_KE[b] != null) {
                        if (jamKe >= jamAktual) break;
                        jamKe++;
                    } else {
                        boolean adaJamBerikut = false;
                        for (int nb = b+1; nb < barisPer; nb++) {
                            if (JAM_KE[nb] != null && jamKe < jamAktual) { adaJamBerikut = true; break; }
                        }
                        if (!adaJamBerikut) break;
                    }

                    Row row = sheet.createRow(rowIdx++);
                    row.setHeightInPoints(14);

                    if (LBL_KHUSUS[b] != null) {
                        guruCell(wb, row, 0, "",               fN, CYAN, HorizontalAlignment.CENTER, true, true);
                        guruCell(wb, row, 1, "",               fN, CYAN, HorizontalAlignment.CENTER, true, true);
                        guruCell(wb, row, 2, WAKTU_BARIS[b],   fN, CYAN, HorizontalAlignment.CENTER, true, true);
                        guruCell(wb, row, 3, LBL_KHUSUS[b],    fB, CYAN, HorizontalAlignment.CENTER, true, true);
                        sheet.addMergedRegion(new CellRangeAddress(rowIdx-1, rowIdx-1, 3, 4));
                        guruCell(wb, row, 4, "",               fN, CYAN, HorizontalAlignment.CENTER, true, true);
                    } else {
                        int slot = start + jamKe - 1;
                        String kelas = "";
                        String mapel = "";
                        for (int k = 0; k < allKelas.length; k++) {
                            if (slot < jadwal.length && k < jadwal[slot].length) {
                                String isi = jadwal[slot][k];
                                if (isi != null && !isi.isEmpty() && isi.replaceAll("[^0-9]","").equals(guruNum)) {
                                    kelas = allKelas[k];
                                    mapel = guruKelasMapel.getOrDefault(guruNum + "|" + k, "");
                                    break;
                                }
                            }
                        }
                        guruCell(wb, row, 0, "",        fN, WHITE, HorizontalAlignment.CENTER, true, true);
                        guruCell(wb, row, 1, JAM_KE[b], fN, WHITE, HorizontalAlignment.CENTER, true, true);
                        guruCell(wb, row, 2, WAKTU_BARIS[b], fN, WHITE, HorizontalAlignment.CENTER, true, true);
                        guruCell(wb, row, 3, kelas,    fN, WHITE, HorizontalAlignment.CENTER, true, true);
                        guruCell(wb, row, 4, mapel,    fN, WHITE, HorizontalAlignment.LEFT, true, true);
                    }
                }
                if (startRow < rowIdx) {
                    sheet.addMergedRegion(new CellRangeAddress(startRow, rowIdx-1, 0, 0));
                    sheet.getRow(startRow).getCell(0).setCellValue(NAMA_HARI[h]);
                    sheet.getRow(startRow).getCell(0).setCellStyle(guruStyle(wb, fB, LGRAY, HorizontalAlignment.CENTER, true, true));
                }
                sheet.createRow(rowIdx++).setHeightInPoints(6);
            }
            try (FileOutputStream fos = new FileOutputStream(outputPath)) { wb.write(fos); }
        }
    }

    static void exportSingleKelasToFile(int kelasIdx, String[][] jadwal, int[][] hariRange,
                                        List<String[]> kebutuhan, String outputPath) throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            final String WHITE = "FFFFFF";
            final String LGRAY = "F2F2F2";
            final String CYAN  = "B2EBF2";

            final String[] WAKTU_BARIS = {
                "06.30 - 07.15","07.15 - 07.55","07.55 - 08.35","08.35 - 09.15","09.15 - 09.55",
                "09.55 - 10.25","10.25 - 11.05","11.05 - 11.45","11.45 - 12.20",
                "12.20 - 13.00","13.00 - 13.40","13.40 - 14.20","14.20 - 15.00"
            };
            final String[] JAM_KE     = {null,"1","2","3","4",null,"5","6",null,"7","8","9","10"};
            final String[] LBL_KHUSUS = {"Sholat Dhuha",null,null,null,null,"Istirahat",null,null,"Sholat Zuhur",null,null,null,null};
            final String[] NAMA_HARI  = {"SENIN","SELASA","RABU","KAMIS","JUMAT"};
            final int barisPer = WAKTU_BARIS.length;

            String[] allKelas = buatDaftarKelas();
            String namaKelas = allKelas[kelasIdx];

            Map<String, String> guruNama = new HashMap<>();
            Map<String, String> guruKelasMapel = new HashMap<>();
            for (String[] data : kebutuhan) {
                String num = data[0].replaceAll("[^0-9]","");
                guruNama.putIfAbsent(num, data[1]);
                try {
                    int kIdx = Integer.parseInt(data[3]);
                    guruKelasMapel.put(num+"|"+kIdx, data[2]);
                } catch (Exception ignored) {}
            }

            String sheetName = "Kelas " + namaKelas;
            XSSFSheet sheet = wb.createSheet(sheetName);
            sheet.setDefaultRowHeightInPoints(14);

            XSSFFont fN = guruFont(wb, 10, false);
            XSSFFont fB = guruFont(wb, 10, true);
            int[] cw = {8,5,16,25,25};
            for (int i = 0; i < cw.length; i++) sheet.setColumnWidth(i, cw[i]*256);

            int rowIdx = 0;
            Row rK = sheet.createRow(rowIdx++); rK.setHeightInPoints(20);
            guruCell(wb, rK, 0, "KELAS", fB, WHITE, HorizontalAlignment.LEFT, false, false);
            guruCell(wb, rK, 1, ":",      fB, WHITE, HorizontalAlignment.CENTER, false, false);
            guruCell(wb, rK, 2, namaKelas, fB, WHITE, HorizontalAlignment.LEFT, false, false);

            rowIdx++;
            Row rH = sheet.createRow(rowIdx++); rH.setHeightInPoints(16);
            guruCell(wb, rH, 0, "HARI",      fB, LGRAY, HorizontalAlignment.CENTER, true, true);
            guruCell(wb, rH, 1, "JAM",       fB, LGRAY, HorizontalAlignment.CENTER, true, true);
            guruCell(wb, rH, 2, "WAKTU",     fB, LGRAY, HorizontalAlignment.CENTER, true, true);
            guruCell(wb, rH, 3, "MAPEL",     fB, LGRAY, HorizontalAlignment.CENTER, true, true);
            guruCell(wb, rH, 4, "NAMA GURU", fB, LGRAY, HorizontalAlignment.CENTER, true, true);

            for (int h = 0; h < hariRange.length; h++) {
                int start = hariRange[h][0], end = hariRange[h][1];
                int jamAktual = end - start + 1;
                int startRow = rowIdx;
                int jamKe = 0;

                for (int b = 0; b < barisPer; b++) {
                    if (JAM_KE[b] != null) {
                        if (jamKe >= jamAktual) break;
                        jamKe++;
                    } else {
                        boolean adaJamBerikut = false;
                        for (int nb = b+1; nb < barisPer; nb++) {
                            if (JAM_KE[nb] != null && jamKe < jamAktual) { adaJamBerikut = true; break; }
                        }
                        if (!adaJamBerikut) break;
                    }

                    Row row = sheet.createRow(rowIdx++);
                    row.setHeightInPoints(14);

                    if (LBL_KHUSUS[b] != null) {
                        guruCell(wb, row, 0, "",               fN, CYAN, HorizontalAlignment.CENTER, true, true);
                        guruCell(wb, row, 1, "",               fN, CYAN, HorizontalAlignment.CENTER, true, true);
                        guruCell(wb, row, 2, WAKTU_BARIS[b],   fN, CYAN, HorizontalAlignment.CENTER, true, true);
                        guruCell(wb, row, 3, LBL_KHUSUS[b],    fB, CYAN, HorizontalAlignment.CENTER, true, true);
                        sheet.addMergedRegion(new CellRangeAddress(rowIdx-1, rowIdx-1, 3, 4));
                        guruCell(wb, row, 4, "",               fN, CYAN, HorizontalAlignment.CENTER, true, true);
                    } else {
                        int slot = start + jamKe - 1;
                        String guruId = "";
                        String mapel = "";
                        String guruName = "";
                        if (slot < jadwal.length && kelasIdx < jadwal[slot].length) {
                            String isi = jadwal[slot][kelasIdx];
                            if (isi != null && !isi.isEmpty()) {
                                guruId = isi.replaceAll("[^0-9]","");
                                guruName = guruNama.getOrDefault(guruId, "");
                                mapel = guruKelasMapel.getOrDefault(guruId+"|"+kelasIdx, "");
                            }
                        }
                        guruCell(wb, row, 0, "",        fN, WHITE, HorizontalAlignment.CENTER, true, true);
                        guruCell(wb, row, 1, JAM_KE[b], fN, WHITE, HorizontalAlignment.CENTER, true, true);
                        guruCell(wb, row, 2, WAKTU_BARIS[b], fN, WHITE, HorizontalAlignment.CENTER, true, true);
                        guruCell(wb, row, 3, mapel,    fN, WHITE, HorizontalAlignment.CENTER, true, true);
                        guruCell(wb, row, 4, guruName, fN, WHITE, HorizontalAlignment.LEFT, true, true);
                    }
                }
                if (startRow < rowIdx) {
                    sheet.addMergedRegion(new CellRangeAddress(startRow, rowIdx-1, 0, 0));
                    sheet.getRow(startRow).getCell(0).setCellValue(NAMA_HARI[h]);
                    sheet.getRow(startRow).getCell(0).setCellStyle(guruStyle(wb, fB, LGRAY, HorizontalAlignment.CENTER, true, true));
                }
                sheet.createRow(rowIdx++).setHeightInPoints(6);
            }
            try (FileOutputStream fos = new FileOutputStream(outputPath)) { wb.write(fos); }
        }
    }

    // =======================================================================
// HELPER BERSAMA (guru + kelas)
// =======================================================================
    static String[] buatDaftarKelas() {
        int n7 = SchedulerUI.kelasT7;
        int n8 = SchedulerUI.kelasT8;
        int n9 = SchedulerUI.kelasT9;
        String[] result = new String[n7+n8+n9];
        int idx = 0;
        for (int i = 0; i < n7; i++) result[idx++] = "7"+(char)('A'+i);
        for (int i = 0; i < n8; i++) result[idx++] = "8"+(char)('A'+i);
        for (int i = 0; i < n9; i++) result[idx++] = "9"+(char)('A'+i);
        return result;
    }

    private static XSSFFont guruFont(XSSFWorkbook wb, int size, boolean bold) {
        XSSFFont f = wb.createFont();
        f.setFontName("Arial");
        f.setFontHeightInPoints((short) size);
        f.setBold(bold);
        return f;
    }

    private static XSSFCellStyle guruStyle(XSSFWorkbook wb, XSSFFont font, String bgHex,
                                           HorizontalAlignment ha, boolean wrap, boolean border) {
        XSSFCellStyle s = wb.createCellStyle();
        s.setFont(font);
        s.setAlignment(ha);
        s.setVerticalAlignment(VerticalAlignment.CENTER);
        s.setWrapText(wrap);
        if (bgHex != null && !bgHex.equals("FFFFFF")) {
            s.setFillForegroundColor(new XSSFColor(hexToRgb(bgHex), null));
            s.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }
        if (border) {
            s.setBorderTop(BorderStyle.THIN);    s.setBorderBottom(BorderStyle.THIN);
            s.setBorderLeft(BorderStyle.THIN);   s.setBorderRight(BorderStyle.THIN);
        }
        return s;
    }

    private static void guruCell(XSSFWorkbook wb, Row row, int col, String value,
                                 XSSFFont font, String bgHex, HorizontalAlignment ha,
                                 boolean wrap, boolean border) {
        Cell c = row.getCell(col);
        if (c == null) c = row.createCell(col);
        c.setCellValue(value);
        c.setCellStyle(guruStyle(wb, font, bgHex, ha, wrap, border));
    }


    private static CellStyle createStyle(Workbook wb, Font font, String bgHex,
                                         HorizontalAlignment ha, boolean wrap, boolean border) {
        CellStyle style = wb.createCellStyle();
        style.setFont(font);
        style.setAlignment(ha);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setWrapText(wrap);

        if (bgHex != null && !bgHex.equals("FFFFFF")) {
            XSSFCellStyle xs = (XSSFCellStyle) style;
            xs.setFillForegroundColor(new XSSFColor(hexToRgb(bgHex), null));
            xs.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }

        if (border) {
            style.setBorderTop(BorderStyle.THIN);
            style.setBorderBottom(BorderStyle.THIN);
            style.setBorderLeft(BorderStyle.THIN);
            style.setBorderRight(BorderStyle.THIN);
        }
        return style;
    }

    private static byte[] hexToRgb(String hex) {
        int r = Integer.parseInt(hex.substring(0, 2), 16);
        int g = Integer.parseInt(hex.substring(2, 4), 16);
        int b = Integer.parseInt(hex.substring(4, 6), 16);
        return new byte[]{(byte) r, (byte) g, (byte) b};
    }


    // Hitung bentrok (tidak berubah)
    public static int hitungBentrok(String[][] jadwal) {
        int penalti = 0;
        for (int i = 0; i < jadwal.length; i++) {
            for (int j = 0; j < jadwal[i].length; j++) {
                String a = jadwal[i][j];
                if (a == null || a.equals("")) continue;
                String angkaA = a.replaceAll("[^0-9]", "");
                for (int k = j + 1; k < jadwal[i].length; k++) {
                    String b = jadwal[i][k];
                    if (b == null || b.equals("")) continue;
                    String angkaB = b.replaceAll("[^0-9]", "");
                    if (angkaA.equals(angkaB)) penalti++;
                }
            }
        }
        return penalti;
    }

    static int hitungPenaltiPJOK(String[][] jadwal, int[][] hariRange, Set<String> guruPJOK) {
        int penaltiTotal = 0;
        for (int h = 0; h < hariRange.length; h++) {
            int start = hariRange[h][0];
            int end   = hariRange[h][1];
            for (int r = start; r <= end; r++) {
                int jam = r - start + 1;
                // Cek pasangan (r, r+1): jam ke-r dan jam ke-(r+1)
                // Pasangan melanggar kalau jam ke-(r+1) > 5, yaitu jam >= 5
                if (jam < 5) continue;           // ← ubah dari <= 5 ke < 5
                if (r + 1 > end) continue;

                for (int k = 0; k < jadwal[r].length; k++) {
                    String idGuru = jadwal[r][k];
                    if (idGuru == null || idGuru.isEmpty()) continue;
                    if (!guruPJOK.contains(idGuru)) continue;

                    String nextGuru = jadwal[r + 1][k];
                    if (idGuru.equals(nextGuru)) {
                        penaltiTotal++;
                    }
                }
            }
        }
        return penaltiTotal;
    }

    static int hitungPenaltiMGMPSenin(String[][] jadwal,
                                      int[][] hariRange,
                                      Set<String> MGMPsenin) {

        int penaltiTotal = 0;

        for (int r = 0; r < jadwal.length; r++) {

            int hari = getHari(r, hariRange);

            if (hari != 0) continue; // hanya Senin

            int start = hariRange[0][0];
            int jam = r - start + 1;

            // Lewati jam 1 - 4
            if (jam <= 4) continue;

            for (int k = 0; k < jadwal[r].length; k++) {

                String idAsli = jadwal[r][k];

                if (idAsli == null || idAsli.equals("")) continue;

                // ambil hanya angka
                String idGuru = idAsli.replaceAll("[^0-9]", "");

                if (MGMPsenin.contains(idGuru)) {
                    penaltiTotal++;
                }
            }
        }

        return penaltiTotal;
    }

    static int hitungPenaltiMGMPSelasa(String[][] jadwal,
                                       int[][] hariRange,
                                       Set<String> MGMPselasa) {

        int penaltiTotal = 0;

        for (int r = 0; r < jadwal.length; r++) {

            int hari = getHari(r, hariRange);

            if (hari != 1) continue; // hanya Senin

            int start = hariRange[1][0];
            int jam = r - start + 1;

            // Lewati jam 1 - 4
            if (jam <= 4) continue;

            for (int k = 0; k < jadwal[r].length; k++) {

                String idAsli = jadwal[r][k];

                if (idAsli == null || idAsli.equals("")) continue;

                // ambil hanya angka
                String idGuru = idAsli.replaceAll("[^0-9]", "");

                if (MGMPselasa.contains(idGuru)) {
                    penaltiTotal++;
                }
            }
        }

        return penaltiTotal;
    }

    static int hitungPenaltiMGMPRabu(String[][] jadwal,
                                     int[][] hariRange,
                                     Set<String> MGMPrabu) {

        int penaltiTotal = 0;

        for (int r = 0; r < jadwal.length; r++) {

            int hari = getHari(r, hariRange);

            if (hari != 2) continue; // hanya rabu

            int start = hariRange[2][0];
            int jam = r - start + 1;

            // Lewati jam 1 - 4
            if (jam <= 4) continue;

            for (int k = 0; k < jadwal[r].length; k++) {

                String idAsli = jadwal[r][k];

                if (idAsli == null || idAsli.equals("")) continue;

                // ambil hanya angka
                String idGuru = idAsli.replaceAll("[^0-9]", "");

                if (MGMPrabu.contains(idGuru)) {
                    penaltiTotal++;
                }
            }
        }

        return penaltiTotal;
    }

    static int hitungPenaltiMGMPKamis(String[][] jadwal,
                                      int[][] hariRange,
                                      Set<String> MGMPkamis) {

        int penaltiTotal = 0;

        for (int r = 0; r < jadwal.length; r++) {

            int hari = getHari(r, hariRange);

            if (hari != 3) continue; // hanya kamis

            int start = hariRange[3][0];
            int jam = r - start + 1;

            // Lewati jam 1 - 4
            if (jam <= 4) continue;

            for (int k = 0; k < jadwal[r].length; k++) {

                String idAsli = jadwal[r][k];

                if (idAsli == null || idAsli.equals("")) continue;

                // ambil hanya angka
                String idGuru = idAsli.replaceAll("[^0-9]", "");

                if (MGMPkamis.contains(idGuru)) {
                    penaltiTotal++;
                }
            }
        }

        return penaltiTotal;
    }

    static int hitungPenaltiMGMPSeninHard(String[][] jadwal,
                                      int[][] hariRange,
                                      Set<String> MGMPsenin) {

        int penaltiTotal = 0;

        int start = hariRange[0][0];

        // hitung per guru berapa jam mengajar setelah jam 4 di hari Senin
        Map<String, Integer> jamPerGuru = new HashMap<>();

        for (int r = 0; r < jadwal.length; r++) {

            int hari = getHari(r, hariRange);
            if (hari != 0) continue; // hanya Senin

            int jam = r - start + 1;
            if (jam <= 4) continue; // lewati jam 1-4

            for (int k = 0; k < jadwal[r].length; k++) {

                String idAsli = jadwal[r][k];
                if (idAsli == null || idAsli.equals("")) continue;

                String idGuru = idAsli.replaceAll("[^0-9]", "");
                if (!MGMPsenin.contains(idGuru)) continue;

                jamPerGuru.put(idGuru, jamPerGuru.getOrDefault(idGuru, 0) + 1);
            }
        }

        // penalti hanya kalau > 1 jam
        for (int jam : jamPerGuru.values()) {
            if (jam > 1) {
                penaltiTotal += (jam - 1);
            }
        }

        return penaltiTotal;
    }

    static int hitungPenaltiMGMPSelasaHard(String[][] jadwal,
                                       int[][] hariRange,
                                       Set<String> MGMPselasa) {

        int penaltiTotal = 0;

        int start = hariRange[1][0];

        Map<String, Integer> jamPerGuru = new HashMap<>();

        for (int r = 0; r < jadwal.length; r++) {

            int hari = getHari(r, hariRange);
            if (hari != 1) continue; // hanya Selasa

            int jam = r - start + 1;
            if (jam <= 4) continue;

            for (int k = 0; k < jadwal[r].length; k++) {

                String idAsli = jadwal[r][k];
                if (idAsli == null || idAsli.equals("")) continue;

                String idGuru = idAsli.replaceAll("[^0-9]", "");
                if (!MGMPselasa.contains(idGuru)) continue;

                jamPerGuru.put(idGuru, jamPerGuru.getOrDefault(idGuru, 0) + 1);
            }
        }

        for (int jam : jamPerGuru.values()) {
            if (jam > 1) {
                penaltiTotal += (jam - 1);
            }
        }

        return penaltiTotal;
    }

    static int hitungPenaltiMGMPRabuHard(String[][] jadwal,
                                     int[][] hariRange,
                                     Set<String> MGMPrabu) {

        int penaltiTotal = 0;

        int start = hariRange[2][0];

        Map<String, Integer> jamPerGuru = new HashMap<>();

        for (int r = 0; r < jadwal.length; r++) {

            int hari = getHari(r, hariRange);
            if (hari != 2) continue; // hanya Rabu

            int jam = r - start + 1;
            if (jam <= 4) continue;

            for (int k = 0; k < jadwal[r].length; k++) {

                String idAsli = jadwal[r][k];
                if (idAsli == null || idAsli.equals("")) continue;

                String idGuru = idAsli.replaceAll("[^0-9]", "");
                if (!MGMPrabu.contains(idGuru)) continue;

                jamPerGuru.put(idGuru, jamPerGuru.getOrDefault(idGuru, 0) + 1);
            }
        }

        for (int jam : jamPerGuru.values()) {
            if (jam > 1) {
                penaltiTotal += (jam - 1);
            }
        }

        return penaltiTotal;
    }


    static int hitungPenaltiMGMPKamisHard(String[][] jadwal,
                                      int[][] hariRange,
                                      Set<String> MGMPkamis) {

        int penaltiTotal = 0;

        int start = hariRange[3][0];

        Map<String, Integer> jamPerGuru = new HashMap<>();

        for (int r = 0; r < jadwal.length; r++) {

            int hari = getHari(r, hariRange);
            if (hari != 3) continue; // hanya Kamis

            int jam = r - start + 1;
            if (jam <= 4) continue; // lewati jam 1-4

            for (int k = 0; k < jadwal[r].length; k++) {

                String idAsli = jadwal[r][k];
                if (idAsli == null || idAsli.equals("")) continue;

                String idGuru = idAsli.replaceAll("[^0-9]", "");
                if (!MGMPkamis.contains(idGuru)) continue;

                jamPerGuru.put(idGuru, jamPerGuru.getOrDefault(idGuru, 0) + 1);
            }
        }

        // penalti hanya kalau > 1 jam
        for (int jam : jamPerGuru.values()) {
            if (jam > 1) {
                penaltiTotal += (jam - 1);
            }
        }

        return penaltiTotal;
    }

    static int hitungTotalPenaltiMGMPHard(String[][] jadwal,
                                      int[][] hariRange,
                                      Set<String> MGMPsenin,
                                      Set<String> MGMPselasa,
                                      Set<String> MGMPrabu,
                                      Set<String> MGMPkamis) {

        int total = 0;

        total += hitungPenaltiMGMPSeninHard(jadwal, hariRange, MGMPsenin);
        total += hitungPenaltiMGMPSelasaHard(jadwal, hariRange, MGMPselasa);
        total += hitungPenaltiMGMPRabuHard(jadwal, hariRange, MGMPrabu);
        total += hitungPenaltiMGMPKamisHard(jadwal, hariRange, MGMPkamis)*100;

//        System.out.println(
//                "Penalti MGMP: " +
//                        "Senin=" + hitungPenaltiMGMPSenin(jadwal, hariRange, MGMPsenin) + ", " +
//                        "Selasa=" + hitungPenaltiMGMPSelasa(jadwal, hariRange, MGMPselasa) + ", " +
//                        "Rabu=" + hitungPenaltiMGMPRabu(jadwal, hariRange, MGMPrabu) + ", " +
//                        "Kamis=" + hitungPenaltiMGMPKamis(jadwal, hariRange, MGMPkamis)
//        );

        return total;
    }



    static int hitungTotalPenaltiMGMP(String[][] jadwal,
                                      int[][] hariRange,
                                      Set<String> MGMPsenin,
                                      Set<String> MGMPselasa,
                                      Set<String> MGMPrabu,
                                      Set<String> MGMPkamis) {

        int total = 0;

        total += hitungPenaltiMGMPSenin(jadwal, hariRange, MGMPsenin);
        total += hitungPenaltiMGMPSelasa(jadwal, hariRange, MGMPselasa);
        total += hitungPenaltiMGMPRabu(jadwal, hariRange, MGMPrabu);
        total += hitungPenaltiMGMPKamis(jadwal, hariRange, MGMPkamis);

//        System.out.println(
//                "Penalti MGMP: " +
//                        "Senin=" + hitungPenaltiMGMPSenin(jadwal, hariRange, MGMPsenin) + ", " +
//                        "Selasa=" + hitungPenaltiMGMPSelasa(jadwal, hariRange, MGMPselasa) + ", " +
//                        "Rabu=" + hitungPenaltiMGMPRabu(jadwal, hariRange, MGMPrabu) + ", " +
//                        "Kamis=" + hitungPenaltiMGMPKamis(jadwal, hariRange, MGMPkamis)
//        );

        return total;
    }

    static int hitungPenaltiMatematika(String[][] jadwal,
                                       int[][] hariRange,
                                       Set<String> guruMatematika) {

        int penaltiTotal = 0;

        for (int h = 0; h < hariRange.length; h++) {

            int start = hariRange[h][0];
            int end = hariRange[h][1];

            for (int r = start; r <= end; r++) {

                int jam = r - start + 1;

                // hanya hitung setelah jam ke-8
                if (jam <= 8) continue;

                for (int k = 0; k < jadwal[r].length; k++) {

                    String idGuru = jadwal[r][k];

                    if (idGuru == null || idGuru.equals("")) continue;

                    if (guruMatematika.contains(idGuru)) {
                        penaltiTotal++;
                    }
                }
            }
        }

        return penaltiTotal;
    }

    static int hitungPenaltiJam9dan10(String[][] jadwal,
                                      int[][] hariRange,
                                      int maxMingguan) {

        int penalti = 0;

        HashMap<String, Integer> jumlahGuruMingguan = new HashMap<>();

        // Senin - Kamis
        for (int h = 0; h <= 3; h++) {

            int start = hariRange[h][0];

            int r9 = start + 8;
            int r10 = start + 9;

            // Jam 9
            for (int k = 0; k < jadwal[r9].length; k++) {

                String idAsli = jadwal[r9][k];
                if (idAsli == null || idAsli.equals("")) continue;

                String idGuru = idAsli.replaceAll("[^0-9]", "");

                jumlahGuruMingguan.put(
                        idGuru,
                        jumlahGuruMingguan.getOrDefault(idGuru, 0) + 1
                );
            }

            // Jam 10
            for (int k = 0; k < jadwal[r10].length; k++) {

                String idAsli = jadwal[r10][k];
                if (idAsli == null || idAsli.equals("")) continue;

                String idGuru = idAsli.replaceAll("[^0-9]", "");

                jumlahGuruMingguan.put(
                        idGuru,
                        jumlahGuruMingguan.getOrDefault(idGuru, 0) + 1
                );
            }
        }

        // Hitung penalti mingguan
        for (String id : jumlahGuruMingguan.keySet()) {

            int jumlah = jumlahGuruMingguan.get(id);

            if (jumlah > maxMingguan) {
                penalti += (jumlah - maxMingguan);
            }
        }

        return penalti;
    }

    static int hitungPenaltiMaxJamPerHari(String[][] jadwal,
                                          int[][] hariRange,
                                          int maxJam,
                                          int maxHariPenuh) {
        int penalti = 0;

        // kumpulkan jam per guru per hari
        HashMap<String, Integer> hariPenuhPerGuru = new HashMap<>(); // hitung hari yang = maxJam

        for (int h = 0; h < hariRange.length; h++) {

            int start = hariRange[h][0];
            int end = hariRange[h][1];

            HashMap<String, Integer> jumlahGuru = new HashMap<>();

            // ambil semua jam dalam 1 hari
            for (int r = start; r <= end; r++) {
                for (int k = 0; k < jadwal[r].length; k++) {

                    String idAsli = jadwal[r][k];
                    if (idAsli == null || idAsli.equals("")) continue;

                    String idGuru = idAsli.replaceAll("[^0-9]", "");
                    jumlahGuru.put(idGuru, jumlahGuru.getOrDefault(idGuru, 0) + 1);
                }
            }

            // cek penalti per hari
            for (String id : jumlahGuru.keySet()) {
                int jumlah = jumlahGuru.get(id);

                // constraint 1: per hari maksimal 8 jam
                if (jumlah > maxJam) {
                    penalti += (jumlah - maxJam);
                }

                // hitung berapa hari yang penuh (= maxJam)
                if (jumlah >= maxJam) {
                    hariPenuhPerGuru.put(id, hariPenuhPerGuru.getOrDefault(id, 0) + 1);
                }
            }
        }

        // constraint 2: maksimal 2 hari yang penuh 8 jam
        for (String id : hariPenuhPerGuru.keySet()) {
            int hariPenuh = hariPenuhPerGuru.get(id);
            if (hariPenuh > maxHariPenuh) {
                penalti += (hariPenuh - maxHariPenuh);
            }
        }

        return penalti;
    }

    static int hitungPenaltiMaxGuruPerHari(String[][] jadwal,
                                           int[][] hariRange,
                                           int maxGuruPerHari) {

        int penalti = 0;

        int jumlahKelas = jadwal[0].length;

        // Loop setiap kelas
        for (int k = 0; k < jumlahKelas; k++) {

            // Loop setiap hari
            for (int h = 0; h < hariRange.length; h++) {

                int start = hariRange[h][0];
                int end = hariRange[h][1];

                // Menyimpan guru unik dalam 1 hari di kelas tertentu
                HashSet<String> guruUnik = new HashSet<>();

                for (int r = start; r <= end; r++) {

                    String idAsli = jadwal[r][k];

                    if (idAsli == null || idAsli.equals("")) continue;

                    // Ambil ID angka saja
                    String idGuru = idAsli.replaceAll("[^0-9]", "");

                    guruUnik.add(idGuru);
                }

                // Hitung jumlah guru berbeda
                int jumlahGuru = guruUnik.size();

                // Penalti jika lebih dari batas
                if (jumlahGuru > maxGuruPerHari) {
                    penalti += (jumlahGuru - maxGuruPerHari);

                    // Debug optional
//                    System.out.println("Kelas " + k +
//                            " hari " + h +
//                            " memiliki " + jumlahGuru +
//                            " guru (maks " + maxGuruPerHari + ")");
                }
            }
        }

        return penalti;
    }

    static int hitungPenaltiGuruID43(String[][] jadwal,
                                     int[][] hariRange) {

        int penaltiTotal = 0;

        for (int r = 0; r < jadwal.length; r++) {

            int hari = getHari(r, hariRange);

            // 2 = Rabu, 4 = Jumat
            if (hari != 2 && hari != 4) continue;

            for (int k = 0; k < jadwal[r].length; k++) {

                String idAsli = jadwal[r][k];

                if (idAsli == null || idAsli.equals("")) continue;

                // ambil angka saja
                String idGuru = idAsli.replaceAll("[^0-9]", "");

                // cek apakah guru 43
                if (idGuru.equals("43")) {
                    penaltiTotal++;
                }
            }
        }
        return penaltiTotal;
    }

    static int hitungTotalSemuaPenalti(
            String[][] jadwal,
            int[][] hariRange,
            Set<String> guruPJOK,
            Set<String> MGMPsenin,
            Set<String> MGMPselasa,
            Set<String> MGMPrabu,
            Set<String> MGMPkamis,
            Set<String> guruMatematika
    ) {

        int total = 0;
        total += hitungTotalSemuaPenaltiHardConstrain(jadwal, hariRange, guruPJOK);
        total += hitungTotalSemuaPenaltiSoftConstrain(jadwal, hariRange, MGMPsenin, MGMPselasa, MGMPrabu, MGMPkamis, guruMatematika);

        return total;
    }

    static int hitungTotalSemuaPenaltiHardConstrain(
            String[][] jadwal,
            int[][] hariRange,
            Set<String> guruPJOK
    ) {

        int total = 0;

        total += hitungBentrok(jadwal)*100;
        total += hitungPenaltiPJOK(jadwal, hariRange, guruPJOK);
        total += hitungPenaltiJam9dan10(jadwal, hariRange, 5);
        total += hitungPenaltiMaxJamPerHari(jadwal, hariRange, 8, 2)*50;
        total += hitungPenaltiGuruID43(jadwal, hariRange);
        total += hitungPenaltiMaxGuruPerHari(jadwal, hariRange, 6);


        return total;
    }

    static int hitungTotalSemuaPenaltiSoftConstrain(
            String[][] jadwal,
            int[][] hariRange,
            Set<String> MGMPsenin,
            Set<String> MGMPselasa,
            Set<String> MGMPrabu,
            Set<String> MGMPkamis,
            Set<String> guruMatematika
    ) {

        int total = 0;

        total += hitungTotalPenaltiMGMP(jadwal, hariRange, MGMPsenin, MGMPselasa, MGMPrabu, MGMPkamis);
        total += hitungPenaltiMatematika(jadwal, hariRange, guruMatematika);
//
//        );

        return total;
    }

    static void swap3Random(String[][] jadwal, int[][] hariRange) {

        Random rand = new Random();
        int maxPercobaan = 100;
        int percobaan = 0;

        while (percobaan < maxPercobaan) {

            percobaan++;

            int kelas = rand.nextInt(jadwal[0].length);
            int r = rand.nextInt(jadwal.length);
            int s = rand.nextInt(jadwal.length);

            String guruA = jadwal[r][kelas];
            String guruB = jadwal[s][kelas];

            if (guruA.equals("") || guruB.equals("")) continue;

            // ===== cari awal blok A =====
            int startA = -1;

            if (r + 2 < jadwal.length &&
                    jadwal[r][kelas].equals(jadwal[r + 1][kelas]) &&
                    jadwal[r][kelas].equals(jadwal[r + 2][kelas]) &&
                    getHari(r, hariRange) == getHari(r + 1, hariRange) &&
                    getHari(r, hariRange) == getHari(r + 2, hariRange)) {
                startA = r;

            } else if (r - 1 >= 0 && r + 1 < jadwal.length &&
                    jadwal[r][kelas].equals(jadwal[r - 1][kelas]) &&
                    jadwal[r][kelas].equals(jadwal[r + 1][kelas]) &&
                    getHari(r, hariRange) == getHari(r - 1, hariRange) &&
                    getHari(r, hariRange) == getHari(r + 1, hariRange)) {
                startA = r - 1;

            } else if (r - 2 >= 0 &&
                    jadwal[r][kelas].equals(jadwal[r - 1][kelas]) &&
                    jadwal[r][kelas].equals(jadwal[r - 2][kelas]) &&
                    getHari(r, hariRange) == getHari(r - 1, hariRange) &&
                    getHari(r, hariRange) == getHari(r - 2, hariRange)) {
                startA = r - 2;

            } else {
                continue;
            }

            // ===== cari awal blok B =====
            int startB = -1;

            if (s + 2 < jadwal.length &&
                    jadwal[s][kelas].equals(jadwal[s + 1][kelas]) &&
                    jadwal[s][kelas].equals(jadwal[s + 2][kelas]) &&
                    getHari(s, hariRange) == getHari(s + 1, hariRange) &&
                    getHari(s, hariRange) == getHari(s + 2, hariRange)) {
                startB = s;

            } else if (s - 1 >= 0 && s + 1 < jadwal.length &&
                    jadwal[s][kelas].equals(jadwal[s - 1][kelas]) &&
                    jadwal[s][kelas].equals(jadwal[s + 1][kelas]) &&
                    getHari(s, hariRange) == getHari(s - 1, hariRange) &&
                    getHari(s, hariRange) == getHari(s + 1, hariRange)) {
                startB = s - 1;

            } else if (s - 2 >= 0 &&
                    jadwal[s][kelas].equals(jadwal[s - 1][kelas]) &&
                    jadwal[s][kelas].equals(jadwal[s - 2][kelas]) &&
                    getHari(s, hariRange) == getHari(s - 1, hariRange) &&
                    getHari(s, hariRange) == getHari(s - 2, hariRange)) {
                startB = s - 2;

            } else {
                continue;
            }

            if (startA == startB) continue;
            if (guruA.equals(guruB)) continue;

            int hariA = getHari(startA, hariRange);
            int hariB = getHari(startB, hariRange);

            if (guruSudahAdaDiHari(jadwal, kelas, guruA, hariB, hariRange)) continue;
            if (guruSudahAdaDiHari(jadwal, kelas, guruB, hariA, hariRange)) continue;

            // ===== swap blok 3 =====
            jadwal[startA][kelas]     = guruB;
            jadwal[startA + 1][kelas] = guruB;
            jadwal[startA + 2][kelas] = guruB;

            jadwal[startB][kelas]     = guruA;
            jadwal[startB + 1][kelas] = guruA;
            jadwal[startB + 2][kelas] = guruA;

            return;
        }

//        System.out.println("[swap3Random] gagal setelah " + maxPercobaan + " percobaan");
    }

    static void swap2PJOK(String[][] jadwal, int[][] hariRange, Set<String> guruPJOK) {

        Random rand = new Random();
        int iterasi = 0;

        while (iterasi < 1) {

            int kelas = rand.nextInt(jadwal[0].length);

            int r = rand.nextInt(jadwal.length);
            int s = rand.nextInt(jadwal.length);

            String guruA = jadwal[r][kelas];
            String guruB = jadwal[s][kelas];

            if (guruA.equals("") || guruB.equals("")) continue;

            // salah satu harus PJOK
            boolean aIsPJOK = guruPJOK.contains(guruA);
            boolean bIsPJOK = guruPJOK.contains(guruB);
            if (!aIsPJOK && !bIsPJOK) continue;

            // pastikan yang PJOK ada di r, yang bukan di s
            if (bIsPJOK && !aIsPJOK) {
                // tukar r dan s agar A selalu PJOK
                int tmp = r; r = s; s = tmp;
                String tmpG = guruA; guruA = guruB; guruB = tmpG;
            }

            // ===== cari awal blok A =====
            int startA = -1;

            if (r < jadwal.length - 1 &&
                    jadwal[r][kelas].equals(jadwal[r + 1][kelas]) &&
                    getHari(r, hariRange) == getHari(r + 1, hariRange)) {
                startA = r;
            } else if (r > 0 &&
                    jadwal[r][kelas].equals(jadwal[r - 1][kelas]) &&
                    getHari(r, hariRange) == getHari(r - 1, hariRange)) {
                startA = r - 1;
            } else {
                continue;
            }

            // ===== cari awal blok B =====
            int startB = -1;

            if (s < jadwal.length - 1 &&
                    jadwal[s][kelas].equals(jadwal[s + 1][kelas]) &&
                    getHari(s, hariRange) == getHari(s + 1, hariRange)) {
                startB = s;
            } else if (s > 0 &&
                    jadwal[s][kelas].equals(jadwal[s - 1][kelas]) &&
                    getHari(s, hariRange) == getHari(s - 1, hariRange)) {
                startB = s - 1;
            } else {
                continue;
            }

            if (startA == startB) continue;
            if (Math.abs(startA - startB) <= 1) continue;

            // ===== cek blok A bukan bagian dari 3 berurutan =====
            boolean aIsBlok3 = false;
            if (startA - 1 >= 0 &&
                    jadwal[startA - 1][kelas].equals(guruA) &&
                    getHari(startA - 1, hariRange) == getHari(startA, hariRange)) {
                aIsBlok3 = true;
            }
            if (startA + 2 < jadwal.length &&
                    jadwal[startA + 2][kelas].equals(guruA) &&
                    getHari(startA + 2, hariRange) == getHari(startA, hariRange)) {
                aIsBlok3 = true;
            }
            if (aIsBlok3) continue;

            // ===== cek blok B bukan bagian dari 3 berurutan =====
            boolean bIsBlok3 = false;
            if (startB - 1 >= 0 &&
                    jadwal[startB - 1][kelas].equals(guruB) &&
                    getHari(startB - 1, hariRange) == getHari(startB, hariRange)) {
                bIsBlok3 = true;
            }
            if (startB + 2 < jadwal.length &&
                    jadwal[startB + 2][kelas].equals(guruB) &&
                    getHari(startB + 2, hariRange) == getHari(startB, hariRange)) {
                bIsBlok3 = true;
            }
            if (bIsBlok3) continue;

            if (guruA.equals(guruB)) continue;

            int hariA = getHari(startA, hariRange);
            int hariB = getHari(startB, hariRange);

            // ===== blok B tidak boleh keluar dari jam ke-4 =====
            int startHariB = hariRange[hariB][0];
            int posisiB = startB - startHariB; // 0-based
            if (posisiB > 3) continue; // jam ke-4 ke atas tidak boleh

            if (guruSudahAdaDiHari(jadwal, kelas, guruA, hariB, hariRange)) continue;
            if (guruSudahAdaDiHari(jadwal, kelas, guruB, hariA, hariRange)) continue;

            // ===== swap blok =====
            jadwal[startA][kelas] = guruB;
            jadwal[startA + 1][kelas] = guruB;

            jadwal[startB][kelas] = guruA;
            jadwal[startB + 1][kelas] = guruA;

            iterasi++;
        }
    }

    static void swap2Random(String[][] jadwal, int[][] hariRange) {

        Random rand = new Random();
        int iterasi = 0;

        while (iterasi < 1) {

            int kelas = rand.nextInt(jadwal[0].length);

            int r = rand.nextInt(jadwal.length);
            int s = rand.nextInt(jadwal.length);

            String guruA = jadwal[r][kelas];
            String guruB = jadwal[s][kelas];

            if (guruA.equals("") || guruB.equals("")) continue;

            // ===== cari awal blok A =====
            int startA = -1;

            if (r < jadwal.length - 1 &&
                    jadwal[r][kelas].equals(jadwal[r + 1][kelas]) &&
                    getHari(r, hariRange) == getHari(r + 1, hariRange)) {

                startA = r;

            } else if (r > 0 &&
                    jadwal[r][kelas].equals(jadwal[r - 1][kelas]) &&
                    getHari(r, hariRange) == getHari(r - 1, hariRange)) {

                startA = r - 1;

            } else {
                continue;
            }

            // ===== cari awal blok B =====
            int startB = -1;

            if (s < jadwal.length - 1 &&
                    jadwal[s][kelas].equals(jadwal[s + 1][kelas]) &&
                    getHari(s, hariRange) == getHari(s + 1, hariRange)) {

                startB = s;

            } else if (s > 0 &&
                    jadwal[s][kelas].equals(jadwal[s - 1][kelas]) &&
                    getHari(s, hariRange) == getHari(s - 1, hariRange)) {

                startB = s - 1;

            } else {
                continue;
            }

            if (startA == startB) continue;
            if (Math.abs(startA - startB) <= 1) continue;
            // ===== cek blok A bukan bagian dari 3 berurutan =====
            boolean aIsBlok3 = false;
            if (startA - 1 >= 0 &&
                    jadwal[startA - 1][kelas].equals(guruA) &&
                    getHari(startA - 1, hariRange) == getHari(startA, hariRange)) {
                aIsBlok3 = true;
            }
            if (startA + 2 < jadwal.length &&
                    jadwal[startA + 2][kelas].equals(guruA) &&
                    getHari(startA + 2, hariRange) == getHari(startA, hariRange)) {
                aIsBlok3 = true;
            }
            if (aIsBlok3) continue;

            // ===== cek blok B bukan bagian dari 3 berurutan =====
            boolean bIsBlok3 = false;
            if (startB - 1 >= 0 &&
                    jadwal[startB - 1][kelas].equals(guruB) &&
                    getHari(startB - 1, hariRange) == getHari(startB, hariRange)) {
                bIsBlok3 = true;
            }
            if (startB + 2 < jadwal.length &&
                    jadwal[startB + 2][kelas].equals(guruB) &&
                    getHari(startB + 2, hariRange) == getHari(startB, hariRange)) {
                bIsBlok3 = true;
            }
            if (bIsBlok3) continue;

            if (guruA.equals(guruB)) continue;

            int hariA = getHari(startA, hariRange);
            int hariB = getHari(startB, hariRange);

            if (guruSudahAdaDiHari(jadwal, kelas, guruA, hariB, hariRange)) continue;
            if (guruSudahAdaDiHari(jadwal, kelas, guruB, hariA, hariRange)) continue;

            // ===== swap blok =====
            jadwal[startA][kelas] = guruB;
            jadwal[startA + 1][kelas] = guruB;

            jadwal[startB][kelas] = guruA;
            jadwal[startB + 1][kelas] = guruA;

            iterasi++;

            //System.out.println("Swap2 ke-" + iterasi);
        }
    }

    static void swap1Random(String[][] jadwal, int[][] hariRange) {

        Random rand = new Random();
        int iterasi = 0;

        while (iterasi < 1) {

            int kelas = rand.nextInt(jadwal[0].length);

            int r = rand.nextInt(jadwal.length);
            int s = rand.nextInt(jadwal.length);

            String guruA = jadwal[r][kelas];
            String guruB = jadwal[s][kelas];

            if (guruA.equals("") || guruB.equals("")) continue;

            // ===== cek slot A tunggal =====
            boolean tunggalA = true;

            if (r > 0 &&
                    jadwal[r - 1][kelas].equals(guruA) &&
                    getHari(r - 1, hariRange) == getHari(r, hariRange)) {
                tunggalA = false;
            }

            if (r < jadwal.length - 1 &&
                    jadwal[r + 1][kelas].equals(guruA) &&
                    getHari(r + 1, hariRange) == getHari(r, hariRange)) {
                tunggalA = false;
            }

            if (!tunggalA) continue;

            // ===== cek slot B tunggal =====
            boolean tunggalB = true;

            if (s > 0 &&
                    jadwal[s - 1][kelas].equals(guruB) &&
                    getHari(s - 1, hariRange) == getHari(s, hariRange)) {
                tunggalB = false;
            }

            if (s < jadwal.length - 1 &&
                    jadwal[s + 1][kelas].equals(guruB) &&
                    getHari(s + 1, hariRange) == getHari(s, hariRange)) {
                tunggalB = false;
            }

            if (!tunggalB) continue;

            if (r == s) continue;
            if (guruA.equals(guruB)) continue;

            int hariA = getHari(r, hariRange);
            int hariB = getHari(s, hariRange);

            // ===== cek guruA pindah ke hari B =====
            if (guruSudahAdaDiHari(jadwal, kelas, guruA, hariB, hariRange)) continue;

            // ===== cek guruB pindah ke hari A =====
            if (guruSudahAdaDiHari(jadwal, kelas, guruB, hariA, hariRange)) continue;

//            System.out.println(
//                    "Swap1 ke-" + (iterasi+1) +
//                            " | kelas=" + kelas +
//                            " | rowA=" + r +
//                            " guru=" + guruA +
//                            " | rowB=" + s +
//                            " guru=" + guruB
//            );

            // ===== swap 1 slot =====
            jadwal[r][kelas] = guruB;
            jadwal[s][kelas] = guruA;

            iterasi++;

//            System.out.println("Swap1 berhasil ke-" + iterasi);
        }
    }

    static void swap2WithSingle(String[][] jadwal, int[][] hariRange) {

        Random rand = new Random();
        int maxPercobaan = 100;
        int percobaan = 0;

        while (percobaan < maxPercobaan) {

            percobaan++;

            int kelas = rand.nextInt(jadwal[0].length);

            // =========================
            // STEP 1: PILIH BLOK 2 JAM
            // =========================
            int r = rand.nextInt(jadwal.length);

            String guruBlock = jadwal[r][kelas];

            if (guruBlock.equals("")) continue;

            int startBlock = -1;

            // cek blok ke bawah
            if (r < jadwal.length - 1 &&
                    jadwal[r][kelas].equals(jadwal[r + 1][kelas]) &&
                    getHari(r, hariRange) == getHari(r + 1, hariRange)) {

                startBlock = r;

            }
            // cek blok ke atas
            else if (r > 0 &&
                    jadwal[r][kelas].equals(jadwal[r - 1][kelas]) &&
                    getHari(r, hariRange) == getHari(r - 1, hariRange)) {

                startBlock = r - 1;

            } else {
                continue;
            }

            int hariBlock = getHari(startBlock, hariRange);

            // ===== cek startBlock bukan bagian dari 3 berurutan =====
            boolean blockIsBlok3 = false;
            if (startBlock - 1 >= 0 &&
                    jadwal[startBlock - 1][kelas].equals(guruBlock) &&
                    getHari(startBlock - 1, hariRange) == getHari(startBlock, hariRange)) {
                blockIsBlok3 = true;
            }
            if (startBlock + 2 < jadwal.length &&
                    jadwal[startBlock + 2][kelas].equals(guruBlock) &&
                    getHari(startBlock + 2, hariRange) == getHari(startBlock, hariRange)) {
                blockIsBlok3 = true;
            }
            if (blockIsBlok3) continue;

            // =========================
            // STEP 2: PILIH DUA SLOT SINGLE BERURUTAN
            // =========================
            List<Integer> indexGanjil = new ArrayList<>();
            for (int i = 0; i < jadwal.length - 1; i++) {
                int hari = getHari(i, hariRange);
                int jamKe = i - hariRange[hari][0] + 1;

                boolean s1Ganjil = (jamKe % 2 == 1);
                boolean s2MasihSatuHari = getHari(i + 1, hariRange) == hari;

                if (s1Ganjil && s2MasihSatuHari) {
                    indexGanjil.add(i);
                }
            }

            if (indexGanjil.isEmpty()) continue;

            int s1 = indexGanjil.get(rand.nextInt(indexGanjil.size()));
            int s2 = s1 + 1;

            // harus dalam hari yang sama
            if (getHari(s1, hariRange) != getHari(s2, hariRange)) continue;

            String guruSingle1 = jadwal[s1][kelas];
            String guruSingle2 = jadwal[s2][kelas];

            if (guruSingle1.equals("") || guruSingle2.equals("")) continue;

            // tidak boleh overlap dengan blok utama
            if (s1 == startBlock || s1 == startBlock + 1 ||
                    s2 == startBlock || s2 == startBlock + 1) {
                continue;
            }

            // =========================
            // STEP 3: CEK SINGLE MURNI
            // =========================
            // cek s1: hanya lihat luar pasangan (jangan cek ke s2)
            boolean s1IsBlock =
                    (s1 - 1 >= 0 &&
                            jadwal[s1][kelas].equals(jadwal[s1 - 1][kelas]) &&
                            getHari(s1, hariRange) == getHari(s1 - 1, hariRange))
                            ||
                            // cek kanan tapi bukan s2
                            (s2 + 1 < jadwal.length &&
                                    jadwal[s1][kelas].equals(jadwal[s2 + 1][kelas]) &&
                                    getHari(s1, hariRange) == getHari(s2 + 1, hariRange));

            // cek s2: hanya lihat luar pasangan (jangan cek ke s1)
            boolean s2IsBlock =
                    // cek kiri tapi bukan s1
                    (s1 - 1 >= 0 &&
                            jadwal[s2][kelas].equals(jadwal[s1 - 1][kelas]) &&
                            getHari(s2, hariRange) == getHari(s1 - 1, hariRange))
                            ||
                            (s2 + 1 < jadwal.length &&
                                    jadwal[s2][kelas].equals(jadwal[s2 + 1][kelas]) &&
                                    getHari(s2, hariRange) == getHari(s2 + 1, hariRange));

            if (s1IsBlock || s2IsBlock) continue;

            int hariSingle = getHari(s1, hariRange);

            if (guruSudahAdaDiHari(jadwal, kelas, guruBlock, hariSingle, hariRange)) continue;

            if (guruSudahAdaDiHari(jadwal, kelas, guruSingle1, hariBlock, hariRange)) continue;
            if (guruSudahAdaDiHari(jadwal, kelas, guruSingle2, hariBlock, hariRange)) continue;

            // hindari guru sama
            if (guruBlock.equals(guruSingle1) ||
                    guruBlock.equals(guruSingle2)) {
                continue;
            }

            // =========================
            // STEP 4: LAKUKAN SWAP
            // =========================
            jadwal[startBlock][kelas]     = guruSingle1;
            jadwal[startBlock + 1][kelas] = guruSingle2;

            jadwal[s1][kelas] = guruBlock;
            jadwal[s2][kelas] = guruBlock;

            return;
        }

//        if (percobaan >= maxPercobaan) {
//            System.out.println("[swap2WithSingle] gagal setelah " + maxPercobaan + " percobaan");
//        }
    }

    static void swap3WithDoubleAndSingle(String[][] jadwal, int[][] hariRange) {

        Random rand = new Random();

        int maxPercobaan = 100;
        int percobaan = 0;

        while (percobaan < maxPercobaan) {

            percobaan++;

            int kelas = rand.nextInt(jadwal[0].length);

            // =========================
            // STEP 1: PILIH BLOK-3
            // =========================
            int r = rand.nextInt(jadwal.length);

            String guruBlock3 = jadwal[r][kelas];
            if (guruBlock3.equals("")) continue;

            int startBlock3 = -1;

            if (r + 2 < jadwal.length &&
                    jadwal[r][kelas].equals(jadwal[r + 1][kelas]) &&
                    jadwal[r][kelas].equals(jadwal[r + 2][kelas]) &&
                    getHari(r, hariRange) == getHari(r + 1, hariRange) &&
                    getHari(r, hariRange) == getHari(r + 2, hariRange)) {
                startBlock3 = r;

            } else if (r - 1 >= 0 && r + 1 < jadwal.length &&
                    jadwal[r][kelas].equals(jadwal[r - 1][kelas]) &&
                    jadwal[r][kelas].equals(jadwal[r + 1][kelas]) &&
                    getHari(r, hariRange) == getHari(r - 1, hariRange) &&
                    getHari(r, hariRange) == getHari(r + 1, hariRange)) {
                startBlock3 = r - 1;

            } else if (r - 2 >= 0 &&
                    jadwal[r][kelas].equals(jadwal[r - 1][kelas]) &&
                    jadwal[r][kelas].equals(jadwal[r - 2][kelas]) &&
                    getHari(r, hariRange) == getHari(r - 1, hariRange) &&
                    getHari(r, hariRange) == getHari(r - 2, hariRange)) {
                startBlock3 = r - 2;

            } else {
                continue;
            }

            // Hitung posisi dan tentukan susunan yang dicari
            int startHariBlock3 = hariRange[getHari(startBlock3, hariRange)][0];
            int posisiDalamHari = (startBlock3 - startHariBlock3) + 1;
            boolean startGanjil = (posisiDalamHari % 2 != 0);

            // cek istirahat
            int posisiBlok2DalamHari = startGanjil ? posisiDalamHari : posisiDalamHari + 1;
            if (posisiBlok2DalamHari == 4 || posisiBlok2DalamHari == 6) continue;

            int hariBlock3 = getHari(startBlock3, hariRange);

            // =========================
            // STEP 2: CARI PASANGAN DOUBLE+SINGLE atau SINGLE+DOUBLE
            // berurutan 3 slot dalam 1 hari
            // startGanjil → cari [double][double][single]
            // startGenap  → cari [single][double][double]
            // =========================
            List<Integer> kandidatPasangan = new ArrayList<>();

            for (int i = 0; i < jadwal.length - 2; i++) {

                // harus 3 slot dalam hari yang sama
                if (getHari(i, hariRange) != getHari(i + 1, hariRange)) continue;
                if (getHari(i, hariRange) != getHari(i + 2, hariRange)) continue;

                int hariI = getHari(i, hariRange);

                String g0 = jadwal[i][kelas];
                String g1 = jadwal[i + 1][kelas];
                String g2 = jadwal[i + 2][kelas];

                if (g0.equals("") || g1.equals("") || g2.equals("")) continue;

                // tidak boleh sama dengan guruBlock3
                if (g0.equals(guruBlock3) || g1.equals(guruBlock3) || g2.equals(guruBlock3)) continue;

                if (startGanjil) {
                    // cari [double][double][single]
                    // g0 == g1 (double), g2 beda (single)
                    if (!g0.equals(g1)) continue;
                    if (g2.equals(g0)) continue; // g2 harus beda dari double

                    String guruDouble = g0;
                    String guruSingle = g2;
                    if (guruDouble.equals(guruSingle)) continue;

                    // cek double murni: kiri i-1 dan kanan i+2(=g2 sudah beda) bukan bagian blok-3
                    boolean doubleKiriAman = (i - 1 < 0 ||
                            !jadwal[i - 1][kelas].equals(guruDouble) ||
                            getHari(i - 1, hariRange) != hariI);
                    // kanan double adalah i+2 yang sudah beda guru → aman
                    if (!doubleKiriAman) continue;

                    // cek single murni: kiri (i+1 = guruDouble, beda) → aman
                    // kanan i+3 harus beda
                    boolean singleKananAman = (i + 3 >= jadwal.length ||
                            !jadwal[i + 3][kelas].equals(guruSingle) ||
                            getHari(i + 3, hariRange) != hariI);
                    if (!singleKananAman) continue;

                } else {
                    // cari [single][double][double]
                    // g0 beda (single), g1 == g2 (double)
                    if (!g1.equals(g2)) continue;
                    if (g0.equals(g1)) continue; // g0 harus beda dari double

                    String guruSingle = g0;
                    String guruDouble = g1;
                    if (guruSingle.equals(guruDouble)) continue;

                    // cek single murni: kiri i-1 harus beda
                    boolean singleKiriAman = (i - 1 < 0 ||
                            !jadwal[i - 1][kelas].equals(guruSingle) ||
                            getHari(i - 1, hariRange) != hariI);
                    if (!singleKiriAman) continue;
                    // kanan single adalah i+1 = guruDouble, beda → aman

                    // cek double murni: kiri i(=guruSingle, beda) → aman
                    // kanan i+3 harus beda
                    boolean doubleKananAman = (i + 3 >= jadwal.length ||
                            !jadwal[i + 3][kelas].equals(guruDouble) ||
                            getHari(i + 3, hariRange) != hariI);
                    if (!doubleKananAman) continue;
                }

                // constraint: guruDouble dan guruSingle belum ada di hariBlock3
                String guruDouble = startGanjil ? g0 : g1;
                String guruSingle = startGanjil ? g2 : g0;

                if (guruSudahAdaDiHari(jadwal, kelas, guruDouble, hariBlock3, hariRange)) continue;
                if (guruSudahAdaDiHari(jadwal, kelas, guruSingle, hariBlock3, hariRange)) continue;

                // constraint: guruBlock3 belum ada di hariI
                if (guruSudahAdaDiHari(jadwal, kelas, guruBlock3, hariI, hariRange)) continue;

                kandidatPasangan.add(i);
            }

            if (kandidatPasangan.isEmpty()) continue;

            int startPasangan = kandidatPasangan.get(rand.nextInt(kandidatPasangan.size()));
            String g0 = jadwal[startPasangan][kelas];
            String g1 = jadwal[startPasangan + 1][kelas];
            String g2 = jadwal[startPasangan + 2][kelas];

            String guruDouble = startGanjil ? g0 : g1;
            String guruSingle = startGanjil ? g2 : g0;
            int hariPasangan  = getHari(startPasangan, hariRange);
            int hariJumat = hariRange.length - 1;
            if (hariPasangan == hariJumat) {
                int startHariJumat = hariRange[hariJumat][0];
                int posisiDiJumat = startPasangan - startHariJumat;
                if (posisiDiJumat >= 4) continue; // block jam ke-5 Jumat
            }


            // =========================
            // STEP 3: LAKUKAN SWAP
            // =========================

            // kosongkan blok-3
            jadwal[startBlock3][kelas]     = "";
            jadwal[startBlock3 + 1][kelas] = "";
            jadwal[startBlock3 + 2][kelas] = "";

            // kosongkan pasangan
            jadwal[startPasangan][kelas]     = "";
            jadwal[startPasangan + 1][kelas] = "";
            jadwal[startPasangan + 2][kelas] = "";

            // isi posisi blok-3 dengan pasangan (double+single atau single+double)
            if (startGanjil) {
                // [double][double][single]
                jadwal[startBlock3][kelas]     = guruDouble;
                jadwal[startBlock3 + 1][kelas] = guruDouble;
                jadwal[startBlock3 + 2][kelas] = guruSingle;
            } else {
                // [single][double][double]
                jadwal[startBlock3][kelas]     = guruSingle;
                jadwal[startBlock3 + 1][kelas] = guruDouble;
                jadwal[startBlock3 + 2][kelas] = guruDouble;
            }

            // isi posisi pasangan dengan blok-3
            jadwal[startPasangan][kelas]     = guruBlock3;
            jadwal[startPasangan + 1][kelas] = guruBlock3;
            jadwal[startPasangan + 2][kelas] = guruBlock3;

            return;
        }
    }

    static void swapRotasiDouble(String[][] jadwal, int[][] hariRange) {

        Random rand = new Random();
        int iterasi = 0;

        while (iterasi < 1) {

            int kelas = rand.nextInt(jadwal[0].length);

            // ===== pilih hari random =====
            int hari = rand.nextInt(hariRange.length);
            int startHari = hariRange[hari][0];
            int endHari = hariRange[hari][1];
            int jumlahJam = endHari - startHari + 1;

            // harus genap
            if (jumlahJam % 2 != 0) continue;

            // ===== ambil 2 slot paling atas =====
            String slot1 = jadwal[startHari][kelas];
            String slot2 = jadwal[startHari + 1][kelas];

            // ===== geser semua slot ke atas 2 posisi =====
            for (int r = startHari; r <= endHari - 2; r++) {
                jadwal[r][kelas] = jadwal[r + 2][kelas];
            }

            // ===== taruh 2 slot tadi ke paling bawah =====
            jadwal[endHari - 1][kelas] = slot1;
            jadwal[endHari][kelas] = slot2;

            iterasi++;
        }
    }

    private static void swapHariSatuKelas(String[][] jadwal, int[][] hariRange) {
        Random rand = new Random();

        int kelas = rand.nextInt(jadwal[0].length);

        int hari1 = rand.nextInt(hariRange.length);
        int hari2;
        do {
            hari2 = rand.nextInt(hariRange.length);
        } while (hari2 == hari1);

        int start1 = hariRange[hari1][0];
        int end1   = hariRange[hari1][1];
        int start2 = hariRange[hari2][0];
        int end2   = hariRange[hari2][1];

        // Cek panjang hari harus sama
        int panjang1 = end1 - start1;
        int panjang2 = end2 - start2;
        if (panjang1 != panjang2) return; // skip kalau beda

        // Tukar slot per slot
        for (int offset = 0; offset <= panjang1; offset++) {
            String temp = jadwal[start1 + offset][kelas];
            jadwal[start1 + offset][kelas] = jadwal[start2 + offset][kelas];
            jadwal[start2 + offset][kelas] = temp;
        }
    }


    static String[][] copyJadwal(String[][] jadwal) {
        String[][] copy = new String[jadwal.length][];

        for (int i = 0; i < jadwal.length; i++) {
            copy[i] = jadwal[i].clone();
        }

        return copy;
    }

    static int[] iterasiGlobal = {0};

    static String[][] hillClimbingTabuGuru43(
            String[][] jadwal,
            int[][] hariRange,
            Set<String> guruPJOK) {

        Random rand = new Random();
        Queue<Integer> tabuList = new LinkedList<>();
        int tabuSize = 2;

        String[][] current = copyJadwal(jadwal);

        int penaltiSekarang = hitungPenaltiGuruID43(current, hariRange);
        int penaltiPJOKSekarang = hitungPenaltiPJOK(current, hariRange, guruPJOK);

        int iterasi = 0;

        //System.out.println("Penalti awal: " + penaltiSekarang);

        while (iterasi < 1000000) {
            if (SchedulerUI.stopRequested) break;
            iterasi++;
            iterasiGlobal[0]++;

//            if (iterasiGlobal[0] % 20000 == 0) {
//                System.out.println("[Iterasi " + iterasiGlobal[0] + "] Penalti: " + penaltiSekarang );
//            }

            if (penaltiSekarang <= 0) {
                System.out.println("Pelanggaran Guru ID43 Optimal");
                break;
            }

            String[][] neighbor = copyJadwal(current);

            //tabu
            int move;

            do {
                move = rand.nextInt(4);
            } while (tabuList.contains(move));

            if (move == 0) {
                swap1Random(neighbor, hariRange);
            } else if (move == 1) {
                swap2WithSingle(neighbor, hariRange);
            } else if (move == 2) {
                swap3WithDoubleAndSingle(neighbor, hariRange);
            }else{
                swapHariSatuKelas(neighbor, hariRange);
            }

            tabuList.offer(move);
            if (tabuList.size() > tabuSize) {
                tabuList.poll();
            }

            int penaltiBaru = hitungPenaltiGuruID43(neighbor, hariRange);
            int penaltiPJOKBaru = hitungPenaltiPJOK(neighbor, hariRange, guruPJOK);


//            System.out.println(
//                    "Total: " + penaltiSekarang + " -> " + penaltiBaru +
//                            " | PJOK: " + penaltiPJOKSekarang + " -> " + penaltiPJOKBaru
//            );

            if (penaltiBaru <= penaltiSekarang
                    && penaltiPJOKSekarang >= penaltiPJOKBaru){

                for (int i = 0; i < current.length; i++) {
                    for (int j = 0; j < current[i].length; j++) {
                        current[i][j] = neighbor[i][j];
                    }
                }
                penaltiSekarang = penaltiBaru;
                penaltiPJOKSekarang = penaltiPJOKBaru;

//                System.out.println("Diterima");

            } else {
                for (int i = 0; i < neighbor.length; i++) {
                    for (int j = 0; j < neighbor[i].length; j++) {
                        neighbor[i][j] = current[i][j];
                    }
                }
//                System.out.println("Ditolak");
            }
        }
        System.out.println("Optimasi Pelanggaran Guru ID 43 Berhenti di iterasi global: " + iterasiGlobal[0] + " | Pelanggaran akhir: " + penaltiSekarang);
        return current;
    }

    static String[][] hillClimbingTabuHardConstrain(
            String[][] jadwal,
            int[][] hariRange,
            Set<String> guruPJOK) {

        Random rand = new Random();
        Queue<Integer> tabuList = new LinkedList<>();
        int tabuSize = 2;

        String[][] current = copyJadwal(jadwal);

        int penaltiSekarang = hitungTotalSemuaPenaltiHardConstrain(current, hariRange, guruPJOK);
        int penaltiPJOKSekarang = hitungPenaltiPJOK(current, hariRange, guruPJOK);
        int penaltiGuru43Sekarang = hitungPenaltiGuruID43(current, hariRange);

        int iterasi = 0;

        System.out.println("Pelanggaran Hard Constraint awal: " + penaltiSekarang);

        while (iterasi < 1000000) {
            if (SchedulerUI.stopRequested) break;
            iterasi++;
            iterasiGlobal[0]++;

            if (iterasiGlobal[0] % 20000 == 0) {
                System.out.println("[Iterasi " + iterasiGlobal[0] + "] Pelanggaran: " + penaltiSekarang );
            }

            if (penaltiSekarang <= 0) {
                System.out.println("Hard Constrain Optimal");
                break;
            }

            String[][] neighbor = copyJadwal(current);

            //tabu
            int move;
            do {
                move = rand.nextInt(4);
            } while (tabuList.contains(move));

            if (move == 0) {
                swap1Random(neighbor, hariRange);
            } else if (move == 1) {
                swap2Random(neighbor, hariRange);
            } else if (move == 2) {
                swap2WithSingle(neighbor, hariRange);
            } else {
                swapRotasiDouble(neighbor, hariRange);
            }

            tabuList.offer(move);
            if (tabuList.size() > tabuSize) {
                tabuList.poll();
            }

            int penaltiBaru = hitungTotalSemuaPenaltiHardConstrain(neighbor, hariRange, guruPJOK);
            int penaltiPJOKBaru = hitungPenaltiPJOK(neighbor, hariRange, guruPJOK);
            int penaltiGuru43Baru = hitungPenaltiGuruID43(neighbor, hariRange);


//            System.out.println(
//                    "Total: " + penaltiSekarang + " -> " + penaltiBaru +
//                            " | PJOK: " + penaltiPJOKSekarang + " -> " + penaltiPJOKBaru
//            );

            if (penaltiBaru <= penaltiSekarang
                    && penaltiPJOKSekarang >= penaltiPJOKBaru
                    && penaltiGuru43Sekarang >= penaltiGuru43Baru){

                for (int i = 0; i < current.length; i++) {
                    for (int j = 0; j < current[i].length; j++) {
                        current[i][j] = neighbor[i][j];
                    }
                }
                penaltiSekarang = penaltiBaru;
                penaltiPJOKSekarang = penaltiPJOKBaru;
                penaltiGuru43Sekarang = penaltiGuru43Baru;

//                System.out.println("Diterima");

            } else {
                for (int i = 0; i < neighbor.length; i++) {
                    for (int j = 0; j < neighbor[i].length; j++) {
                        neighbor[i][j] = current[i][j];
                    }
                }
//                System.out.println("Ditolak");
            }
        }
        System.out.println("Optimasi hard constrain berhenti di iterasi global: " + iterasiGlobal[0] + " | Pelanggaran akhir: " + penaltiSekarang);
        return current;
    }

    static String[][] hillClimbingTabuSoftConstrain(
            String[][] jadwal,
            int[][] hariRange,
            Set<String> guruPJOK,
            Set<String> MGMPsenin,
            Set<String> MGMPselasa,
            Set<String> MGMPrabu,
            Set<String> MGMPkamis,
            Set<String> guruMatematika) {

        Random rand = new Random();

        String[][] current = copyJadwal(jadwal);

        Queue<Integer> tabuList = new LinkedList<>();
        int tabuSize = 2;

        int penaltiSekarang = hitungTotalSemuaPenaltiSoftConstrain(current, hariRange, MGMPsenin, MGMPselasa, MGMPrabu, MGMPkamis, guruMatematika);
        int penaltiHardConstrainSekarang = hitungTotalSemuaPenaltiHardConstrain(current, hariRange, guruPJOK);

        int iterasi = 0;

        System.out.println("Penalti awal: " + penaltiSekarang);

        while (iterasi < 1000000) {
            iterasi++;

            if (SchedulerUI.stopRequested) {
                System.out.println("» Stop diminta, keluar dari iterasi.");
                break;
            }
            iterasiGlobal[0]++;

            if (iterasiGlobal[0] % 20000 == 0) {
                System.out.println("[Iterasi " + iterasiGlobal[0] + "] Penalti: " + penaltiSekarang);
            }

            if (penaltiSekarang <= 0) {
                System.out.println("Target tercapai");
                break;
            }

            String[][] neighbor = copyJadwal(current);

            //random


            //tabu

            int move;
            do {
                move = rand.nextInt(4);
            } while (tabuList.contains(move));

            if (move == 0) {
                swap1Random(neighbor, hariRange);
            } else if (move == 1) {
                swap2Random(neighbor, hariRange);
            } else if (move == 2) {
                swap2WithSingle(neighbor, hariRange);
            } else {
                swapRotasiDouble(neighbor, hariRange);
            }

            tabuList.offer(move);
            if (tabuList.size() > tabuSize) {
                tabuList.poll();
            }

            int penaltiBaru = hitungTotalSemuaPenaltiSoftConstrain(neighbor, hariRange, MGMPsenin, MGMPselasa, MGMPrabu, MGMPkamis, guruMatematika);
            int penaltiHardConstrainBaru = hitungTotalSemuaPenaltiHardConstrain(neighbor, hariRange, guruPJOK);

//            System.out.println(
//                    "Total: " + penaltiSekarang + " -> " + penaltiBaru +
//                            " | PJOK: " + penaltiPJOKSekarang + " -> " + penaltiPJOKBaru
//            );

            if (penaltiBaru <= penaltiSekarang
                    && penaltiHardConstrainBaru <= penaltiHardConstrainSekarang){

                for (int i = 0; i < current.length; i++) {
                    for (int j = 0; j < current[i].length; j++) {
                        current[i][j] = neighbor[i][j];
                    }
                }
                penaltiSekarang = penaltiBaru;
                penaltiHardConstrainSekarang = penaltiHardConstrainBaru;

//                System.out.println("Diterima");

            } else {
                for (int i = 0; i < neighbor.length; i++) {
                    for (int j = 0; j < neighbor[i].length; j++) {
                        neighbor[i][j] = current[i][j];
                    }
                }
//                System.out.println("Ditolak");
            }
        }

        System.out.println("Selesai di iterasi global: " + iterasiGlobal[0] + " | Penalti akhir: " + penaltiSekarang);
        return current;
    }

    static String[][] LAHCTabuSoftConstrain(
            String[][] jadwal,
            int[][] hariRange,
            Set<String> guruPJOK,
            Set<String> MGMPsenin,
            Set<String> MGMPselasa,
            Set<String> MGMPrabu,
            Set<String> MGMPkamis,
            Set<String> guruMatematika) {

        Random rand = new Random();

        String[][] current = copyJadwal(jadwal);
        String[][] best = copyJadwal(jadwal);

        Queue<Integer> tabuList = new LinkedList<>();
        int tabuSize = 2;

        // ── LAHC: tambah costList dan pointer v ─────────────────────
        int L = 100;
        int[] costList = new int[L];
        int v = 0;
        // ────────────────────────────────────────────────────────────

        int penaltiSekarang = hitungTotalSemuaPenaltiSoftConstrain(current, hariRange, MGMPsenin, MGMPselasa, MGMPrabu, MGMPkamis, guruMatematika);
        int penaltiHardConstrainSekarang = hitungTotalSemuaPenaltiHardConstrain(current, hariRange, guruPJOK);
        int penaltiBest = penaltiSekarang;


        // ── LAHC: isi costList dengan penalti awal ──────────────────
        Arrays.fill(costList, penaltiSekarang);
        // ────────────────────────────────────────────────────────────

        int iterasi = 0;

        System.out.println("Penalti soft constrain awal: " + penaltiSekarang);

        while (iterasi < 1000000) {
            iterasi++;

            if (SchedulerUI.stopRequested) {
                System.out.println("» Stop diminta, keluar dari iterasi.");
                break;
            }
            iterasiGlobal[0]++;

            if (iterasiGlobal[0] % 20000 == 0) {
                System.out.println("[Iterasi " + iterasiGlobal[0] + "] Pelanggaran: " + penaltiBest);
            }

            if (penaltiSekarang <= 0) {
                System.out.println("Soft Constrain Optimal!");
                break;
            }

            String[][] neighbor = copyJadwal(current);

            int move;
            do {
                move = rand.nextInt(4);
            } while (tabuList.contains(move));

            if (move == 0) {
                swap1Random(neighbor, hariRange);
            } else if (move == 1) {
                swap2Random(neighbor, hariRange);
            } else if (move == 2) {
                swap2WithSingle(neighbor, hariRange);
            } else {
                swapRotasiDouble(neighbor, hariRange);
            }
            tabuList.offer(move);
            if (tabuList.size() > tabuSize) {
                tabuList.poll();
            }

            int penaltiBaru = hitungTotalSemuaPenaltiSoftConstrain(neighbor, hariRange, MGMPsenin, MGMPselasa, MGMPrabu, MGMPkamis, guruMatematika);
            int penaltiHardConstrainBaru = hitungTotalSemuaPenaltiHardConstrain(neighbor, hariRange, guruPJOK);

            // ── Hanya baris ini yang berubah: <= costList[v] ─────────
            if ((penaltiBaru <= penaltiSekarang || penaltiBaru <= costList[v])
                    && penaltiHardConstrainBaru <= penaltiHardConstrainSekarang) {

                for (int i = 0; i < current.length; i++) {
                    for (int j = 0; j < current[i].length; j++) {
                        current[i][j] = neighbor[i][j];
                    }
                }
                penaltiSekarang = penaltiBaru;
                penaltiHardConstrainSekarang = penaltiHardConstrainBaru;

                if (penaltiSekarang <= penaltiBest) {
                    best = copyJadwal(current);
                    penaltiBest = penaltiSekarang;
                }

            } else {
                for (int i = 0; i < neighbor.length; i++) {
                    for (int j = 0; j < neighbor[i].length; j++) {
                        neighbor[i][j] = current[i][j];
                    }
                }
            }

            // ── LAHC: update history dan geser pointer ───────────────
            costList[v] = penaltiSekarang;
            v = (v + 1) % L;
            // ────────────────────────────────────────────────────────
        }
        System.out.println("Optimasi soft constrain berhenti di iterasi global: " + iterasiGlobal[0] + " | Pelanggaran akhir: " + penaltiSekarang);
        return best;
    }


    static String[][] hillClimbingRandomHardConstrain(
            String[][] jadwal,
            int[][] hariRange,
            Set<String> guruPJOK) {

        Random rand = new Random();

        String[][] current = copyJadwal(jadwal);

        int penaltiSekarang = hitungTotalSemuaPenaltiHardConstrain(current, hariRange, guruPJOK);
        int penaltiPJOKSekarang = hitungPenaltiPJOK(current, hariRange, guruPJOK);

        int iterasi = 0;

        System.out.println("Penalti awal: " + penaltiSekarang);

        while (iterasi < 1000000) {
            if (SchedulerUI.stopRequested) break;
            iterasi++;
            iterasiGlobal[0]++;

            if (iterasiGlobal[0] % 20000 == 0) {
                System.out.println("[Iterasi " + iterasiGlobal[0] + "] Penalti: " + penaltiSekarang);
            }

            if (penaltiSekarang <= 0) {
               System.out.println("Hard Constrain Optimal");

                break;
            }

            String[][] neighbor = copyJadwal(current);

            // random
            int move = rand.nextInt(8);
            if (move == 0) {
                swap1Random(neighbor, hariRange);
            } else if (move == 1) {
                swap2Random(neighbor, hariRange);
            } else if (move == 2) {
                swap2WithSingle(neighbor, hariRange);
            } else if (move == 3) {
                swap3WithDoubleAndSingle(neighbor, hariRange);
            }else if (move == 4){
                swap2PJOK(neighbor, hariRange, guruPJOK);
            }else if (move == 5) {
                swapHariSatuKelas(neighbor, hariRange);
            }else if (move == 6){
                    swapRotasiDouble(neighbor, hariRange);
            } else {
                swap3Random(neighbor, hariRange);
            }


            int penaltiBaru = hitungTotalSemuaPenaltiHardConstrain(neighbor, hariRange, guruPJOK);
            int penaltiPJOKBaru = hitungPenaltiPJOK(neighbor, hariRange, guruPJOK);

//            System.out.println(
//                    "Total: " + penaltiSekarang + " -> " + penaltiBaru +
//                            " | PJOK: " + penaltiPJOKSekarang + " -> " + penaltiPJOKBaru
//            );

            if (penaltiBaru <= penaltiSekarang
                    && penaltiPJOKSekarang >= penaltiPJOKBaru){

                for (int i = 0; i < current.length; i++) {
                    for (int j = 0; j < current[i].length; j++) {
                        current[i][j] = neighbor[i][j];
                    }
                }
                penaltiSekarang = penaltiBaru;
                penaltiPJOKSekarang = penaltiPJOKBaru;
//                System.out.println("Diterima");

            } else {
                for (int i = 0; i < neighbor.length; i++) {
                    for (int j = 0; j < neighbor[i].length; j++) {
                        neighbor[i][j] = current[i][j];
                    }
                }
//                System.out.println("Ditolak");
            }
        }
        System.out.println("Selesai di iterasi global: " + iterasiGlobal[0] + " | Penalti akhir: " + penaltiSekarang);

        return current;
    }

    static String[][] hillClimbingRandomSoftConstrain(
            String[][] jadwal,
            int[][] hariRange,
            Set<String> guruPJOK,
            Set<String> MGMPsenin,
            Set<String> MGMPselasa,
            Set<String> MGMPrabu,
            Set<String> MGMPkamis,
            Set<String> guruMatematika) {

        Random rand = new Random();

        String[][] current = copyJadwal(jadwal);

        int penaltiSekarang = hitungTotalSemuaPenaltiSoftConstrain(current, hariRange, MGMPsenin, MGMPselasa, MGMPrabu, MGMPkamis, guruMatematika);
        int penaltiHardConstrainSekarang = hitungTotalSemuaPenaltiHardConstrain(current, hariRange, guruPJOK);

        int iterasi = 0;

        System.out.println("Penalti awal: " + penaltiSekarang);

        while (iterasi < 1000000) {
            iterasi++;

            if (SchedulerUI.stopRequested) {
                System.out.println("» Stop diminta, keluar dari iterasi.");
                break;
            }
            iterasiGlobal[0]++;

            if (iterasiGlobal[0] % 20000 == 0) {
                System.out.println("[Iterasi " + iterasiGlobal[0] + "] Penalti: " + penaltiSekarang);
            }

            if (penaltiSekarang <= 0) {
                System.out.println("Target tercapai");
                break;
            }

            String[][] neighbor = copyJadwal(current);

            //random
            int move = rand.nextInt(7);
            if (move == 0) {
                swap1Random(neighbor, hariRange);
            } else if (move == 1) {
                swap2Random(neighbor, hariRange);
            } else if (move == 2) {
                swap2WithSingle(neighbor, hariRange);
            } else if (move == 3) {
                swap3WithDoubleAndSingle(neighbor, hariRange);
            }else if (move == 4){
                swap2PJOK(neighbor, hariRange, guruPJOK);
            }else if (move == 5){
                swapHariSatuKelas(neighbor, hariRange);
            } else {
                swap3Random(neighbor, hariRange);
            }


            int penaltiBaru = hitungTotalSemuaPenaltiSoftConstrain(neighbor, hariRange, MGMPsenin, MGMPselasa, MGMPrabu, MGMPkamis, guruMatematika);
            int penaltiHardConstrainBaru = hitungTotalSemuaPenaltiHardConstrain(neighbor, hariRange, guruPJOK);

//            System.out.println(
//                    "Total: " + penaltiSekarang + " -> " + penaltiBaru +
//                            " | PJOK: " + penaltiPJOKSekarang + " -> " + penaltiPJOKBaru
//            );

            if (penaltiBaru <= penaltiSekarang
                    && penaltiHardConstrainBaru<= penaltiHardConstrainSekarang){

                for (int i = 0; i < current.length; i++) {
                    for (int j = 0; j < current[i].length; j++) {
                        current[i][j] = neighbor[i][j];
                    }
                }
                penaltiSekarang = penaltiBaru;
                penaltiHardConstrainSekarang = penaltiHardConstrainBaru;

//                System.out.println("Diterima");

            } else {
                for (int i = 0; i < neighbor.length; i++) {
                    for (int j = 0; j < neighbor[i].length; j++) {
                        neighbor[i][j] = current[i][j];
                    }
                }
//                System.out.println("Ditolak");
            }
        }

        System.out.println("Selesai di iterasi global: " + iterasiGlobal[0] + " | Penalti akhir: " + penaltiSekarang);
        return current;
    }


}