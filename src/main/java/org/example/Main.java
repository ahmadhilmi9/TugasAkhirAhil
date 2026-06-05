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
        Sheet sheet = wb.getSheetAt(3);

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

        // print data guru di arraylist kebutuhan
        System.out.println("=== DATA GURU DAN BEBAN ===");
        for (String[] d : kebutuhan) {
            System.out.println(d[0] + " | " + d[1] + " | " + d[2] + " | " + d[3] + " | " + d[4]);
        }
        System.out.println("Total data: " + kebutuhan.size());

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

        //Memulai Initial Solution
        initialSolutionPJOK(kebutuhan, jadwal, hariRange);

        kebutuhan.sort((a, b) -> Integer.parseInt(a[4]) - Integer.parseInt(b[4]));
       //kebutuhan.sort((a, b) -> Integer.parseInt(b[4]) - Integer.parseInt(a[4]));
        initialSolutionBlok3(kebutuhan, jadwal, hariRange);
        initialSolutionBlok2(kebutuhan, jadwal, hariRange);
        initialSolutionSwap2(kebutuhan, jadwal, hariRange, guruPJOK);
        initialSolutionBlok1(kebutuhan, jadwal, hariRange);


        System.out.println("=== DATA GURU DAN BEBAN ===");
        for (String[] d : kebutuhan) {
            System.out.println(d[0] + " | " + d[1] + " | " + d[2] + " | " + d[3] + " | " + d[4]);
        }
        System.out.println("Total data: " + kebutuhan.size());

        System.out.println("Penalti PJOK awal: " + hitungPenaltiPJOK(jadwal, hariRange, guruPJOK));

        long mulai = System.nanoTime();
        //jadwal = hillClimbingRandomMGMPKamis(jadwal, hariRange, guruPJOK, MGMPkamis);
       //jadwal = hillClimbingRandomHardConstrain(jadwal, hariRange, guruPJOK, MGMPkamis);
       //jadwal = hillClimbingRandomSoftConstrain(jadwal, hariRange, guruPJOK, MGMPsenin, MGMPselasa, MGMPrabu, MGMPkamis, guruMatematika);
        // jadwal = LAHCRandomHardConstrain(jadwal, hariRange,guruPJOK);
        System.out.println("Penalti PJOK: " + hitungPenaltiPJOK(jadwal, hariRange, guruPJOK));
        System.out.println("Penalti guru ID 43: " + hitungPenaltiGuruID43(jadwal, hariRange));
        System.out.println("Penalti MGMP Senin: " + hitungPenaltiMGMPSenin(jadwal, hariRange, MGMPsenin));
        System.out.println("Penalti MGMP Selasa: " + hitungPenaltiMGMPSelasa(jadwal, hariRange, MGMPselasa));
        System.out.println("Penalti MGMP Rabu: " + hitungPenaltiMGMPRabu(jadwal, hariRange, MGMPrabu));
        System.out.println("Penalti MGMP Kamis: " + hitungPenaltiMGMPKamis(jadwal, hariRange, MGMPkamis));
        System.out.println("Penalti MGMP awal: " + hitungTotalPenaltiMGMP(jadwal, hariRange, MGMPsenin, MGMPselasa, MGMPrabu, MGMPkamis));
        System.out.println("Penalti Matematika: " + hitungPenaltiMatematika(jadwal, hariRange, guruMatematika));
        System.out.println("Penalti jam 9&10: " + hitungPenaltiJam9dan10(jadwal, hariRange, 5));
        System.out.println("Penalti bentrok: " + hitungBentrok(jadwal));
        System.out.println("Penalti max jam per hari: " + hitungPenaltiMaxJamPerHari(jadwal, hariRange, 8));
        System.out.println("Penalti total awal: " + hitungTotalSemuaPenalti(jadwal,
                hariRange,
                guruPJOK,
                MGMPsenin,
                MGMPselasa,
                MGMPrabu,
                MGMPkamis,
                guruMatematika));
        long end = System.nanoTime();

        double durasiMenit = (end - mulai) / 1_000_000_000.0 / 60.0;
        System.out.println("Waktu eksekusi: " + durasiMenit + " menit");
        exportJadwalToExcel(jadwal, hariRange, SchedulerUI.outputFilePath, kebutuhan);
        SchedulerUI.markDone(true, null);
//       SchedulerUI.markDone(true, "» File berhasil disimpan: " + SchedulerUI.outputFilePath);


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

            for (int i = 0; i < kebutuhan.size(); i++) {

                String[] data = kebutuhan.get(i);

                String idGuru = data[0];
                String mapel  = data[2];
                int kelas     = Integer.parseInt(data[3]);
                int beban     = Integer.parseInt(data[4]);

                if (beban != 3 && beban != 5) continue;
                if (!mapel.equalsIgnoreCase("MATEMATIKA") && !mapel.equalsIgnoreCase("IPA")) continue;


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

        int barisPer = WAKTU_BARIS.length; // 13

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

        // Reset sheet
        int ex = wb.getSheetIndex("Jadwal Guru");
        if (ex >= 0) wb.removeSheetAt(ex);
        XSSFSheet sheet = wb.createSheet("Jadwal Guru");
        sheet.setDefaultRowHeightInPoints(14);

        // Font
        XSSFFont fN = guruFont(wb, 10, false);
        XSSFFont fB = guruFont(wb, 10, true);
        XSSFFont fL = guruFont(wb, 12, true);

        // Lebar kolom: HARI|JAM|WAKTU|KLS|MAPEL|GAP|HARI|JAM|WAKTU|KLS|MAPEL
        int[] cw = {8,4,16,5,20,2,8,4,16,5,20};
        for (int i = 0; i < cw.length; i++) sheet.setColumnWidth(i, cw[i]*256);

        int rowIdx = 0;

        // Judul
        Row rJ = sheet.createRow(rowIdx++); rJ.setHeightInPoints(18);
        guruCell(wb, rJ, 0, "TAHUN AJARAN 2025-2026", fL, WHITE, HorizontalAlignment.CENTER, false, false);
        sheet.addMergedRegion(new CellRangeAddress(0,0,0,10));

        // Info guru (baris 1-3)
        String[] lblInfo = {"KODE GURU","NAMA GURU","JML MENGAJAR"};
        for (int i = 0; i < 3; i++) {
            Row r = sheet.createRow(rowIdx++); r.setHeightInPoints(20);
            guruCell(wb, r, 0, lblInfo[i], fB, WHITE, HorizontalAlignment.LEFT,   false, false);
            guruCell(wb, r, 1, ":",        fB, WHITE, HorizontalAlignment.CENTER, false, false);
            guruCell(wb, r, 2, "",         fL, WHITE, HorizontalAlignment.LEFT,   false, false);
            sheet.addMergedRegion(new CellRangeAddress(rowIdx-1, rowIdx-1, 2, 10));
        }
        // Tombol navigasi
        sheet.getRow(1).createCell(10).setCellValue("▲");
        sheet.getRow(1).getCell(10).setCellStyle(guruStyle(wb, fB, "DDDDDD", HorizontalAlignment.CENTER, false, true));
        sheet.getRow(2).createCell(10).setCellValue("▼");
        sheet.getRow(2).getCell(10).setCellStyle(guruStyle(wb, fB, "DDDDDD", HorizontalAlignment.CENTER, false, true));

        // Header jadwal (2 baris)
        Row rH1 = sheet.createRow(rowIdx++); rH1.setHeightInPoints(14);
        guruCell(wb, rH1, 0, "HARI",          fB, LGRAY, HorizontalAlignment.CENTER, true, true);
        guruCell(wb, rH1, 1, "JADWAL",        fB, LGRAY, HorizontalAlignment.CENTER, true, true);
        sheet.addMergedRegion(new CellRangeAddress(rowIdx-1,rowIdx-1,1,2));
        guruCell(wb, rH1, 3, "KLS",           fB, LGRAY, HorizontalAlignment.CENTER, true, true);
        guruCell(wb, rH1, 4, "MATA PELAJARAN",fB, LGRAY, HorizontalAlignment.CENTER, true, true);
        guruCell(wb, rH1, 5, "",              fB, WHITE, HorizontalAlignment.CENTER, false, false);
        guruCell(wb, rH1, 6, "HARI",          fB, LGRAY, HorizontalAlignment.CENTER, true, true);
        guruCell(wb, rH1, 7, "JADWAL",        fB, LGRAY, HorizontalAlignment.CENTER, true, true);
        sheet.addMergedRegion(new CellRangeAddress(rowIdx-1,rowIdx-1,7,8));
        guruCell(wb, rH1, 9, "KLS",           fB, LGRAY, HorizontalAlignment.CENTER, true, true);
        guruCell(wb, rH1,10, "MATA PELAJARAN",fB, LGRAY, HorizontalAlignment.CENTER, true, true);

        Row rH2 = sheet.createRow(rowIdx++); rH2.setHeightInPoints(14);
        guruCell(wb, rH2, 0, "",      fB, LGRAY, HorizontalAlignment.CENTER, true, true);
        guruCell(wb, rH2, 1, "JAM",  fB, LGRAY, HorizontalAlignment.CENTER, true, true);
        guruCell(wb, rH2, 2, "WAKTU",fB, LGRAY, HorizontalAlignment.CENTER, true, true);
        guruCell(wb, rH2, 3, "",     fB, LGRAY, HorizontalAlignment.CENTER, true, true);
        guruCell(wb, rH2, 4, "",     fB, LGRAY, HorizontalAlignment.CENTER, true, true);
        guruCell(wb, rH2, 5, "",     fB, WHITE, HorizontalAlignment.CENTER, false, false);
        guruCell(wb, rH2, 6, "",     fB, LGRAY, HorizontalAlignment.CENTER, true, true);
        guruCell(wb, rH2, 7, "JAM",  fB, LGRAY, HorizontalAlignment.CENTER, true, true);
        guruCell(wb, rH2, 8, "WAKTU",fB, LGRAY, HorizontalAlignment.CENTER, true, true);
        guruCell(wb, rH2, 9, "",     fB, LGRAY, HorizontalAlignment.CENTER, true, true);
        guruCell(wb, rH2,10, "",     fB, LGRAY, HorizontalAlignment.CENTER, true, true);
        for (int c : new int[]{0,3,4,5,6,9,10})
            sheet.addMergedRegion(new CellRangeAddress(rowIdx-2,rowIdx-1,c,c));

        int dataRowStart = rowIdx;

        // Hitung jam aktual tiap hari
        int[] jamAktualPerHari = new int[hariRange.length];
        for (int h = 0; h < hariRange.length; h++)
            jamAktualPerHari[h] = hariRange[h][1] - hariRange[h][0] + 1;

        int GAP = 2; // baris kosong antar hari

// Helper inline: hitung baris yang dipakai untuk jam aktual tertentu
// Logika: iterasi JAM_KE[], hitung baris sampai jam ke-N terpenuhi
// Baris istirahat/sholat TETAP dihitung selama masih dalam range jam
// --------------------------------------------------------
// Contoh: jam=9 → baris 0(Dhuha) 1 2 3 4 [Ist] 5 6 [SholZhur] 7 8 9 = 11 baris
// Contoh: jam=8 → baris 0(Dhuha) 1 2 3 4 [Ist] 5 6 [SholZhur] 7 8    = 10 baris
// Contoh: jam=10→ semua 13 baris

// Fungsi getBarisPakai sudah ada sebagai helper method terpisah
// Pastikan helper method ini ada:
// private static int getBarisPakai(int jamAktual, String[] JAM_KE, int barisPer)

// =====================
// Hitung posisi start KIRI (Senin=0, Selasa=1, Rabu=2)
// =====================
        int[] kiriStart = new int[3];
        int cursorKiri = rowIdx;
        for (int h = 0; h < 3; h++) {
            kiriStart[h] = cursorKiri;
            cursorKiri += getBarisPakai(jamAktualPerHari[h], JAM_KE, barisPer) + GAP;
        }

// =====================
// Hitung posisi start KANAN (Kamis=3, Jumat=4)
// =====================
        int[] kananStart = new int[2];
        int cursorKanan = rowIdx;
        for (int h = 0; h < 2; h++) {
            kananStart[h] = cursorKanan;
            cursorKanan += getBarisPakai(jamAktualPerHari[h+3], JAM_KE, barisPer) + GAP;
        }

// Total baris keseluruhan
        int totalBarisKiri  = cursorKiri  - GAP - rowIdx; // hapus GAP terakhir
        int totalBarisKanan = cursorKanan - GAP - rowIdx;
        int totalBaris = Math.max(totalBarisKiri, totalBarisKanan);

// Buat semua baris kosong dulu (tanpa border)
        for (int b = 0; b < totalBaris + GAP; b++) {
            Row row = sheet.createRow(rowIdx + b);
            row.setHeightInPoints(14);
            for (int c = 0; c <= 10; c++)
                guruCell(wb, row, c, "", fN, WHITE, HorizontalAlignment.CENTER, false, false);
        }

// =====================
// ISI KIRI: Senin(0), Selasa(1), Rabu(2)
// =====================
        for (int h = 0; h < 3; h++) {
            int base    = kiriStart[h];
            int jam     = jamAktualPerHari[h];
            int barisPakai = getBarisPakai(jam, JAM_KE, barisPer);

            // Merge kolom HARI vertikal
            guruCell(wb, sheet.getRow(base), 0, NAMA_HARI[h], fB, LGRAY,
                    HorizontalAlignment.CENTER, true, true);
            if (barisPakai > 1)
                sheet.addMergedRegion(new CellRangeAddress(base, base+barisPakai-1, 0, 0));

            int jamKe = 0;
            int bRow  = 0; // index baris relatif dalam kotak ini
            for (int b = 0; b < barisPer; b++) {
                // Cek apakah jam sudah habis
                if (JAM_KE[b] != null) {
                    if (jamKe >= jam) break; // stop, jam sudah cukup
                    jamKe++;
                } else {
                    // Baris istirahat/sholat: tampilkan hanya jika jam berikutnya masih ada
                    // Cek apakah masih ada jam setelah baris ini
                    boolean adaJamBerikut = false;
                    for (int nb = b+1; nb < barisPer; nb++) {
                        if (JAM_KE[nb] != null && jamKe < jam) { adaJamBerikut = true; break; }
                    }
                    if (!adaJamBerikut) break; // tidak perlu tampilkan istirahat di akhir
                }

                Row row = sheet.getRow(base + bRow);
                if (LBL_KHUSUS[b] != null) {
                    guruCell(wb, row, 1, "",              fN, CYAN, HorizontalAlignment.CENTER, true, true);
                    guruCell(wb, row, 2, WAKTU_BARIS[b],  fN, CYAN, HorizontalAlignment.CENTER, true, true);
                    guruCell(wb, row, 3, LBL_KHUSUS[b],   fB, CYAN, HorizontalAlignment.CENTER, true, true);
                    guruCell(wb, row, 4, "",              fN, CYAN, HorizontalAlignment.CENTER, true, true);
                    sheet.addMergedRegion(new CellRangeAddress(base+bRow, base+bRow, 3, 4));
                } else {
                    guruCell(wb, row, 1, JAM_KE[b],       fN, WHITE, HorizontalAlignment.CENTER, true, true);
                    guruCell(wb, row, 2, WAKTU_BARIS[b],  fN, WHITE, HorizontalAlignment.CENTER, true, true);
                    guruCell(wb, row, 3, "",              fN, WHITE, HorizontalAlignment.CENTER, true, true);
                    guruCell(wb, row, 4, "",              fN, WHITE, HorizontalAlignment.CENTER, true, true);
                }
                bRow++;
            }
        }

// =====================
// ISI KANAN: Kamis(3), Jumat(4)
// =====================
        for (int h = 0; h < 2; h++) {
            int base    = kananStart[h];
            int jam     = jamAktualPerHari[h+3];
            int barisPakai = getBarisPakai(jam, JAM_KE, barisPer);

            // Merge kolom HARI vertikal
            guruCell(wb, sheet.getRow(base), 6, NAMA_HARI[h+3], fB, LGRAY,
                    HorizontalAlignment.CENTER, true, true);
            if (barisPakai > 1)
                sheet.addMergedRegion(new CellRangeAddress(base, base+barisPakai-1, 6, 6));

            int jamKe = 0;
            int bRow  = 0;
            for (int b = 0; b < barisPer; b++) {
                if (JAM_KE[b] != null) {
                    if (jamKe >= jam) break;
                    jamKe++;
                } else {
                    boolean adaJamBerikut = false;
                    for (int nb = b+1; nb < barisPer; nb++) {
                        if (JAM_KE[nb] != null && jamKe < jam) { adaJamBerikut = true; break; }
                    }
                    if (!adaJamBerikut) break;
                }

                Row row = sheet.getRow(base + bRow);
                if (row == null) row = sheet.createRow(base + bRow);

                if (LBL_KHUSUS[b] != null) {
                    guruCell(wb, row, 7, "",              fN, CYAN, HorizontalAlignment.CENTER, true, true);
                    guruCell(wb, row, 8, WAKTU_BARIS[b],  fN, CYAN, HorizontalAlignment.CENTER, true, true);
                    guruCell(wb, row, 9, LBL_KHUSUS[b],   fB, CYAN, HorizontalAlignment.CENTER, true, true);
                    guruCell(wb, row,10, "",              fN, CYAN, HorizontalAlignment.CENTER, true, true);
                    sheet.addMergedRegion(new CellRangeAddress(base+bRow, base+bRow, 9, 10));
                } else {
                    guruCell(wb, row, 7, JAM_KE[b],       fN, WHITE, HorizontalAlignment.CENTER, true, true);
                    guruCell(wb, row, 8, WAKTU_BARIS[b],  fN, WHITE, HorizontalAlignment.CENTER, true, true);
                    guruCell(wb, row, 9, "",              fN, WHITE, HorizontalAlignment.CENTER, true, true);
                    guruCell(wb, row,10, "",              fN, WHITE, HorizontalAlignment.CENTER, true, true);
                }
                bRow++;
            }
        }
        // Isi border untuk baris 26-38 di kolom kanan (yang tidak terpakai)
        for (int b = barisPer*2; b < barisPer*3; b++) {
            Row row = sheet.getRow(rowIdx + b);
            if (row == null) row = sheet.createRow(rowIdx + b);
            for (int c : new int[]{6,7,8,9,10})
                guruCell(wb, row, c, "", fN, WHITE, HorizontalAlignment.CENTER, true, true);
        }

        // DataGuru sheet tersembunyi
        int exD = wb.getSheetIndex("DataGuru");
        if (exD >= 0) wb.removeSheetAt(exD);
        XSSFSheet ds = wb.createSheet("DataGuru");
        wb.setSheetHidden(wb.getSheetIndex("DataGuru"), true);

        Row dHdr = ds.createRow(0);
        dHdr.createCell(0).setCellValue("KodeGuru");
        dHdr.createCell(1).setCellValue("NamaGuru");
        dHdr.createCell(2).setCellValue("Mapel");
        dHdr.createCell(3).setCellValue("TotalJam");
        int co = 4;
        for (int h = 0; h < hariRange.length; h++)
            for (int b = 0; b < barisPer; b++)
                dHdr.createCell(co++).setCellValue("H"+h+"B"+b);

        int dri = 1;
        for (String guruNum : guruList) {
            String[] info = guruMap.get(guruNum);
            Row dr = ds.createRow(dri++);
            dr.createCell(0).setCellValue(guruNum);
            dr.createCell(1).setCellValue(info[0]);
            dr.createCell(2).setCellValue(info[1]);

            int totalJam = 0, col2 = 4;
            for (int h = 0; h < hariRange.length; h++) {
                int start = hariRange[h][0], end = hariRange[h][1];
                int[] jb = new int[barisPer]; Arrays.fill(jb,-1);
                int jIdx = start;
                for (int b = 0; b < barisPer; b++)
                    if (JAM_KE[b] != null && jIdx <= end) jb[b] = jIdx++;

                for (int b = 0; b < barisPer; b++) {
                    String val = "";
                    if (jb[b] >= 0) {
                        for (int k = 0; k < allKelas.length; k++) {
                            if (k < jadwal[jb[b]].length) {
                                String isi = jadwal[jb[b]][k];
                                if (isi != null && !isi.isEmpty() && isi.replaceAll("[^0-9]","").equals(guruNum)) {
                                    val = allKelas[k]+"|"+info[1];
                                    totalJam++;
                                    break;
                                }
                            }
                        }
                    }
                    dr.createCell(col2++).setCellValue(val);
                }
            }
            dr.createCell(3).setCellValue(totalJam);
        }

        // Isi guru pertama langsung
        if (!guruList.isEmpty()) {
            Row dr = ds.getRow(1);
            sheet.getRow(1).getCell(2).setCellValue(dr.getCell(0).getStringCellValue());
            sheet.getRow(2).getCell(2).setCellValue(dr.getCell(1).getStringCellValue());
            sheet.getRow(3).getCell(2).setCellValue((int)dr.getCell(3).getNumericCellValue()+" JAM");
            isiDataGuruKeSheet(sheet, ds, 1, dataRowStart, barisPer, JAM_KE, LBL_KHUSUS, hariRange);
        }

        // VBACode config
        int vi = wb.getSheetIndex("VBACode");
        if (vi >= 0) wb.removeSheetAt(vi);
        XSSFSheet vs = wb.createSheet("VBACode");
        wb.setSheetHidden(wb.getSheetIndex("VBACode"), true);
        vs.createRow(0).createCell(0).setCellValue("config");
        vs.createRow(1).createCell(0).setCellValue(dataRowStart+1); // startRow 1-indexed
        vs.createRow(2).createCell(0).setCellValue(barisPer*3);     // totalBaris
        vs.createRow(3).createCell(0).setCellValue(barisPer);       // barisPer
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

    private static void isiDataGuruKeSheet(XSSFSheet sheet, XSSFSheet ds, int dsRow,
                                           int dataRowStart, int barisPer,
                                           String[] JAM_KE, String[] LBL_KHUSUS,
                                           int[][] hariRange) { // <-- tambah parameter ini
        Row dr = ds.getRow(dsRow);
        if (dr == null) return;

        for (int h = 0; h < 3; h++) {
            int base = dataRowStart + h*barisPer;
            int jamAktual = hariRange[h][1] - hariRange[h][0] + 1; // <-- hitung jam aktual hari ini

            int jamKe = 0; // <-- counter jam KBM aktual
            for (int b = 0; b < barisPer; b++) {
                if (LBL_KHUSUS[b] != null) continue;
                jamKe++; // <-- tambah tiap ketemu jam KBM

                Cell dc = dr.getCell(4 + h*barisPer + b);
                String val = (dc != null) ? dc.getStringCellValue() : "";

                // Jika jam ini melebihi jam aktual hari → kosongkan
                if (jamKe > jamAktual) val = ""; // <-- ini yang baru

                Row row = sheet.getRow(base+b); if (row == null) continue;
                // Selalu set value (kosong atau isi) supaya baris sebelumnya tidak tersisa
                String[] p = val.isEmpty() ? new String[]{"",""} : val.split("\\|",2);
                row.getCell(3).setCellValue(p[0]);
                row.getCell(4).setCellValue(p.length>1?p[1]:"");
            }
        }

        for (int h = 3; h < 5; h++) {
            int base = dataRowStart + (h-3)*barisPer;
            int jamAktual = hariRange[h][1] - hariRange[h][0] + 1; // <-- hitung jam aktual hari ini

            int jamKe = 0; // <-- counter jam KBM aktual
            for (int b = 0; b < barisPer; b++) {
                if (LBL_KHUSUS[b] != null) continue;
                jamKe++; // <-- tambah tiap ketemu jam KBM

                Cell dc = dr.getCell(4 + h*barisPer + b);
                String val = (dc != null) ? dc.getStringCellValue() : "";

                // Jika jam ini melebihi jam aktual hari → kosongkan
                if (jamKe > jamAktual) val = ""; // <-- ini yang baru

                Row row = sheet.getRow(base+b); if (row == null) continue;
                String[] p = val.isEmpty() ? new String[]{"",""} : val.split("\\|",2);
                row.getCell(9).setCellValue(p[0]);
                row.getCell(10).setCellValue(p.length>1?p[1]:"");
            }
        }
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
        final int[][] PANEL_HARI  = {{0,1},{2,3},{4,-1}};

        int barisPer = WAKTU_BARIS.length;

        List<String> kelasList = new ArrayList<>();
        for (String k : buatDaftarKelas()) kelasList.add(k);
        int totalKelas = kelasList.size();

        // Reset sheet
        int ex = wb.getSheetIndex("Jadwal Kelas");
        if (ex >= 0) wb.removeSheetAt(ex);
        XSSFSheet sheet = wb.createSheet("Jadwal Kelas");
        sheet.setDefaultRowHeightInPoints(14);

        XSSFFont fN = guruFont(wb, 9,  false);
        XSSFFont fB = guruFont(wb, 9,  true);
        XSSFFont fL = guruFont(wb, 11, true);

        int[] cw = {7,4,15,14,22,2,7,4,15,14,22};
        for (int i = 0; i < cw.length; i++) sheet.setColumnWidth(i, cw[i]*256);

        int rowIdx = 0;

        // Judul
        Row rJ1 = sheet.createRow(rowIdx++); rJ1.setHeightInPoints(16);
        guruCell(wb, rJ1, 0, "JADWAL PELAJARAN SEMESTER GANJIL", fL, WHITE, HorizontalAlignment.CENTER, false, false);
        sheet.addMergedRegion(new CellRangeAddress(0,0,0,10));
        Row rJ2 = sheet.createRow(rowIdx++); rJ2.setHeightInPoints(16);
        guruCell(wb, rJ2, 0, "TAHUN AJARAN 2025-2026", fL, WHITE, HorizontalAlignment.CENTER, false, false);
        sheet.addMergedRegion(new CellRangeAddress(1,1,0,10));

        // Info kelas (baris 2-3)
        String[] lblInfo = {"KELAS","WALI KELAS"};
        for (int i = 0; i < 2; i++) {
            Row r = sheet.createRow(rowIdx++); r.setHeightInPoints(20);
            guruCell(wb, r, 0, lblInfo[i], fB, WHITE, HorizontalAlignment.LEFT,   false, false);
            guruCell(wb, r, 1, ":",        fB, WHITE, HorizontalAlignment.CENTER, false, false);
            guruCell(wb, r, 2, "",         fL, WHITE, HorizontalAlignment.LEFT,   false, false);
            sheet.addMergedRegion(new CellRangeAddress(rowIdx-1, rowIdx-1, 2, 10));
        }
        // Tombol navigasi
        sheet.getRow(2).createCell(10).setCellValue("▲");
        sheet.getRow(2).getCell(10).setCellStyle(guruStyle(wb, fB, "DDDDDD", HorizontalAlignment.CENTER, false, true));
        sheet.getRow(3).createCell(10).setCellValue("▼");
        sheet.getRow(3).getCell(10).setCellStyle(guruStyle(wb, fB, "DDDDDD", HorizontalAlignment.CENTER, false, true));

        int dataRowStart = rowIdx;
        int panelH = 2 + barisPer + 1; // header(2) + data(13) + spasi(1)

        // Buat 3 panel
        for (int p = 0; p < 3; p++) {
            int hKiri  = PANEL_HARI[p][0];
            int hKanan = PANEL_HARI[p][1];

            // Header panel
            Row h1 = sheet.createRow(rowIdx++); h1.setHeightInPoints(14);
            guruCell(wb, h1, 0, "HARI",        fB, LGRAY, HorizontalAlignment.CENTER, true, true);
            guruCell(wb, h1, 1, "JADWAL",      fB, LGRAY, HorizontalAlignment.CENTER, true, true);
            sheet.addMergedRegion(new CellRangeAddress(rowIdx-1,rowIdx-1,1,2));
            guruCell(wb, h1, 3, "MAPEL",       fB, LGRAY, HorizontalAlignment.CENTER, true, true);
            guruCell(wb, h1, 4, "NAMA GURU",   fB, LGRAY, HorizontalAlignment.CENTER, true, true);
            guruCell(wb, h1, 5, "",            fB, WHITE, HorizontalAlignment.CENTER, false, false);
            guruCell(wb, h1, 6, "HARI",        fB, LGRAY, HorizontalAlignment.CENTER, true, true);
            guruCell(wb, h1, 7, "JADWAL",      fB, LGRAY, HorizontalAlignment.CENTER, true, true);
            sheet.addMergedRegion(new CellRangeAddress(rowIdx-1,rowIdx-1,7,8));
            guruCell(wb, h1, 9, "MAPEL",       fB, LGRAY, HorizontalAlignment.CENTER, true, true);
            guruCell(wb, h1,10, "NAMA GURU",   fB, LGRAY, HorizontalAlignment.CENTER, true, true);

            Row h2 = sheet.createRow(rowIdx++); h2.setHeightInPoints(14);
            guruCell(wb, h2, 0, "",      fB, LGRAY, HorizontalAlignment.CENTER, true, true);
            guruCell(wb, h2, 1, "JAM",  fB, LGRAY, HorizontalAlignment.CENTER, true, true);
            guruCell(wb, h2, 2, "WAKTU",fB, LGRAY, HorizontalAlignment.CENTER, true, true);
            guruCell(wb, h2, 3, "",     fB, LGRAY, HorizontalAlignment.CENTER, true, true);
            guruCell(wb, h2, 4, "",     fB, LGRAY, HorizontalAlignment.CENTER, true, true);
            guruCell(wb, h2, 5, "",     fB, WHITE, HorizontalAlignment.CENTER, false, false);
            guruCell(wb, h2, 6, "",     fB, LGRAY, HorizontalAlignment.CENTER, true, true);
            guruCell(wb, h2, 7, "JAM",  fB, LGRAY, HorizontalAlignment.CENTER, true, true);
            guruCell(wb, h2, 8, "WAKTU",fB, LGRAY, HorizontalAlignment.CENTER, true, true);
            guruCell(wb, h2, 9, "",     fB, LGRAY, HorizontalAlignment.CENTER, true, true);
            guruCell(wb, h2,10, "",     fB, LGRAY, HorizontalAlignment.CENTER, true, true);
            for (int c : new int[]{0,3,4,5,6,9,10})
                sheet.addMergedRegion(new CellRangeAddress(rowIdx-2,rowIdx-1,c,c));

            int dataStart = rowIdx;

            // 13 baris data
            for (int b = 0; b < barisPer; b++) {
                Row row = sheet.createRow(rowIdx++); row.setHeightInPoints(14);
                guruCell(wb, row, 5, "", fN, WHITE, HorizontalAlignment.CENTER, false, false);
                if (LBL_KHUSUS[b] != null) {
                    guruCell(wb, row, 0, "",            fN, CYAN, HorizontalAlignment.CENTER, false, true);
                    guruCell(wb, row, 1, "",            fN, CYAN, HorizontalAlignment.CENTER, true, true);
                    guruCell(wb, row, 2, WAKTU_BARIS[b],fN, CYAN, HorizontalAlignment.CENTER, true, true);
                    guruCell(wb, row, 3, LBL_KHUSUS[b], fB, CYAN, HorizontalAlignment.CENTER, true, true);
                    guruCell(wb, row, 4, "",            fN, CYAN, HorizontalAlignment.CENTER, true, true);
                    sheet.addMergedRegion(new CellRangeAddress(rowIdx-1,rowIdx-1,3,4));
                    guruCell(wb, row, 6, "",            fN, CYAN, HorizontalAlignment.CENTER, false, true);
                    guruCell(wb, row, 7, "",            fN, CYAN, HorizontalAlignment.CENTER, true, true);
                    guruCell(wb, row, 8, WAKTU_BARIS[b],fN, CYAN, HorizontalAlignment.CENTER, true, true);
                    guruCell(wb, row, 9, LBL_KHUSUS[b], fB, CYAN, HorizontalAlignment.CENTER, true, true);
                    guruCell(wb, row,10, "",            fN, CYAN, HorizontalAlignment.CENTER, true, true);
                    sheet.addMergedRegion(new CellRangeAddress(rowIdx-1,rowIdx-1,9,10));
                } else {
                    guruCell(wb, row, 0, "",                              fN, WHITE, HorizontalAlignment.CENTER, false, true);
                    guruCell(wb, row, 1, JAM_KE[b]!=null?JAM_KE[b]:"",  fN, WHITE, HorizontalAlignment.CENTER, true, true);
                    guruCell(wb, row, 2, WAKTU_BARIS[b],                 fN, WHITE, HorizontalAlignment.CENTER, true, true);
                    guruCell(wb, row, 3, "",                              fN, WHITE, HorizontalAlignment.CENTER, true, true);
                    guruCell(wb, row, 4, "",                              fN, WHITE, HorizontalAlignment.CENTER, true, true);
                    guruCell(wb, row, 6, "",                              fN, WHITE, HorizontalAlignment.CENTER, false, true);
                    guruCell(wb, row, 7, JAM_KE[b]!=null?JAM_KE[b]:"",  fN, WHITE, HorizontalAlignment.CENTER, true, true);
                    guruCell(wb, row, 8, WAKTU_BARIS[b],                 fN, WHITE, HorizontalAlignment.CENTER, true, true);
                    guruCell(wb, row, 9, "",                              fN, WHITE, HorizontalAlignment.CENTER, true, true);
                    guruCell(wb, row,10, "",                              fN, WHITE, HorizontalAlignment.CENTER, true, true);
                }
            }

            // Merge HARI vertikal + isi nama hari
            sheet.addMergedRegion(new CellRangeAddress(dataStart, dataStart+barisPer-1, 0, 0));
            sheet.addMergedRegion(new CellRangeAddress(dataStart, dataStart+barisPer-1, 6, 6));
            sheet.getRow(dataStart).getCell(0).setCellValue(NAMA_HARI[hKiri]);
            sheet.getRow(dataStart).getCell(0).setCellStyle(guruStyle(wb, fB, LGRAY, HorizontalAlignment.CENTER, true, true));
            if (hKanan >= 0) {
                sheet.getRow(dataStart).getCell(6).setCellValue(NAMA_HARI[hKanan]);
                sheet.getRow(dataStart).getCell(6).setCellStyle(guruStyle(wb, fB, LGRAY, HorizontalAlignment.CENTER, true, true));
            }

            // Spasi antar panel
            sheet.createRow(rowIdx++).setHeightInPoints(6);
        }

        // DataKelas sheet tersembunyi
        int exD = wb.getSheetIndex("DataKelas");
        if (exD >= 0) wb.removeSheetAt(exD);
        XSSFSheet dks = wb.createSheet("DataKelas");
        wb.setSheetHidden(wb.getSheetIndex("DataKelas"), true);

        Row dHdr = dks.createRow(0);
        dHdr.createCell(0).setCellValue("NamaKelas");
        dHdr.createCell(1).setCellValue("WaliKelas");
        int co = 2;
        for (int h = 0; h < hariRange.length; h++)
            for (int b = 0; b < barisPer; b++)
                dHdr.createCell(co++).setCellValue("H"+h+"B"+b);

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

        // Isi data per kelas
        String[] allKelas = buatDaftarKelas();
        for (int kIdx = 0; kIdx < totalKelas; kIdx++) {
            Row dr = dks.createRow(kIdx+1);
            dr.createCell(0).setCellValue(kelasList.get(kIdx));
            dr.createCell(1).setCellValue("");

            int col2 = 2;
            for (int h = 0; h < hariRange.length; h++) {
                int start = hariRange[h][0], end = hariRange[h][1];
                int[] jb = new int[barisPer]; Arrays.fill(jb,-1);
                int jIdx = start;
                for (int b = 0; b < barisPer; b++)
                    if (JAM_KE[b] != null && jIdx <= end) jb[b] = jIdx++;

                for (int b = 0; b < barisPer; b++) {
                    String val = "";
                    if (jb[b] >= 0 && jb[b] < jadwal.length && kIdx < jadwal[jb[b]].length) {
                        String isi = jadwal[jb[b]][kIdx];
                        if (isi != null && !isi.isEmpty()) {
                            String guruNum = isi.replaceAll("[^0-9]","");
                            String nama  = guruNama.getOrDefault(guruNum, "");
                            String mapel = guruKelasMapel.getOrDefault(guruNum+"|"+kIdx, "");
                            val = mapel+"|"+nama;
                        }
                    }
                    dr.createCell(col2++).setCellValue(val);
                }
            }
        }

        // Isi kelas pertama langsung
        if (totalKelas > 0) {
            Row dr = dks.getRow(1);
            sheet.getRow(2).getCell(2).setCellValue(dr.getCell(0).getStringCellValue());
            sheet.getRow(3).getCell(2).setCellValue("-");
            isiDataKelasKeSheet(sheet, dks, 1, dataRowStart, barisPer, panelH, LBL_KHUSUS, PANEL_HARI);
        }

        // VBAKelas config
        int vki = wb.getSheetIndex("VBAKelas");
        if (vki >= 0) wb.removeSheetAt(vki);
        XSSFSheet vks = wb.createSheet("VBAKelas");
        wb.setSheetHidden(wb.getSheetIndex("VBAKelas"), true);
        vks.createRow(0).createCell(0).setCellValue("config");
        vks.createRow(1).createCell(0).setCellValue(dataRowStart+1); // startRow 1-indexed
        vks.createRow(2).createCell(0).setCellValue(barisPer);
        vks.createRow(3).createCell(0).setCellValue(panelH);
    }

    private static void isiDataKelasKeSheet(XSSFSheet sheet, XSSFSheet dks, int dkRow,
                                            int dataRowStart, int barisPer, int panelH,
                                            String[] LBL_KHUSUS, int[][] PANEL_HARI) {
        Row dr = dks.getRow(dkRow); if (dr == null) return;
        for (int p = 0; p < 3; p++) {
            int panelDataStart = dataRowStart + p*panelH + 2;
            int hKiri  = PANEL_HARI[p][0];
            int hKanan = PANEL_HARI[p][1];
            for (int b = 0; b < barisPer; b++) {
                if (LBL_KHUSUS[b] != null) continue;
                Row row = sheet.getRow(panelDataStart+b); if (row == null) continue;
                // Kiri
                Cell dc = dr.getCell(2 + hKiri*barisPer + b);
                String val = (dc != null) ? dc.getStringCellValue() : "";
                if (!val.isEmpty()) {
                    String[] pts = val.split("\\|",2);
                    row.getCell(3).setCellValue(pts[0]);
                    row.getCell(4).setCellValue(pts.length>1?pts[1]:"");
                } else {
                    row.getCell(3).setCellValue(""); row.getCell(4).setCellValue("");
                }
                // Kanan
                if (hKanan >= 0) {
                    dc = dr.getCell(2 + hKanan*barisPer + b);
                    val = (dc != null) ? dc.getStringCellValue() : "";
                    if (!val.isEmpty()) {
                        String[] pts = val.split("\\|",2);
                        row.getCell(9).setCellValue(pts[0]);
                        row.getCell(10).setCellValue(pts.length>1?pts[1]:"");
                    } else {
                        row.getCell(9).setCellValue(""); row.getCell(10).setCellValue("");
                    }
                }
            }
        }
    }

    // =======================================================================
// HELPER BERSAMA (guru + kelas)
// =======================================================================
    private static String[] buatDaftarKelas() {
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
                                          int maxJam) {

        int penalti = 0;

        // loop per hari
        for (int h = 0; h < hariRange.length; h++) {

            int start = hariRange[h][0];
            int end = hariRange[h][1];

            HashMap<String, Integer> jumlahGuru = new HashMap<>();

            // ambil semua jam dalam 1 hari
            for (int r = start; r <= end; r++) {

                for (int k = 0; k < jadwal[r].length; k++) {

                    String idAsli = jadwal[r][k];

                    if (idAsli == null || idAsli.equals("")) continue;

                    // ambil angka saja (biar konsisten)
                    String idGuru = idAsli.replaceAll("[^0-9]", "");

                    jumlahGuru.put(idGuru, jumlahGuru.getOrDefault(idGuru, 0) + 1);
                }
            }

            // cek penalti
            for (String id : jumlahGuru.keySet()) {

                int jumlah = jumlahGuru.get(id);

                if (jumlah > maxJam) {
                    penalti += (jumlah - maxJam);
                }
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

        total += hitungBentrok(jadwal);
        total += hitungPenaltiPJOK(jadwal, hariRange, guruPJOK);
        total += hitungPenaltiJam9dan10(jadwal, hariRange, 5);
        total += hitungPenaltiMaxJamPerHari(jadwal, hariRange, 8);
        total += hitungPenaltiGuruID43(jadwal, hariRange);
        total += hitungPenaltiMaxGuruPerHari(jadwal, hariRange, 6);

//        System.out.println(
//                "Rincian penalti: " +
//                        "Bentrok=" + hitungBentrok(jadwal) +
//                        ", PJOK=" + hitungPenaltiPJOK(jadwal, hariRange, guruPJOK) +
//                        ", Jam9&10=" + hitungPenaltiJam9dan10(jadwal, hariRange, 2) +
//                        ", MaxJamPerHari=" + hitungPenaltiMaxJamPerHari(jadwal, hariRange, 8)+
//                        ", GuruID43=" + hitungPenaltiGuruID43(jadwal, hariRange)
//        );

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
        int maxPercobaan = 1;
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

    static void rotationRandom(String[][] jadwal, int[][] hariRange) {

        Random rand = new Random();

        while (true) {

            int kelas = rand.nextInt(jadwal[0].length);
            int hari = rand.nextInt(hariRange.length);

            int start = hariRange[hari][0];
            int end = hariRange[hari][1];

            // cari blok 2 jam
            List<Integer> blokStart = new ArrayList<>();

            for (int i = start; i < end; i++) {

                if (jadwal[i][kelas].equals("")) continue;

                // cek blok 2 jam
                if (i + 1 <= end &&
                        jadwal[i][kelas].equals(jadwal[i + 1][kelas])) {

                    // skip kalau blok 3 (PJOK)
                    if (i + 2 <= end &&
                            jadwal[i][kelas].equals(jadwal[i + 2][kelas])) {
                        continue;
                    }

                    blokStart.add(i);
                    i++; // skip pasangan
                }
            }

            // minimal harus ada 2 blok
            if (blokStart.size() < 2) continue;

            // ===== simpan blok pertama =====
            int first = blokStart.get(0);
            String guruFirst = jadwal[first][kelas];

            // ===== rotation =====
            for (int i = 0; i < blokStart.size() - 1; i++) {
                int curr = blokStart.get(i);
                int next = blokStart.get(i + 1);

                jadwal[curr][kelas] = jadwal[next][kelas];
                jadwal[curr + 1][kelas] = jadwal[next + 1][kelas];
            }

            // ===== taruh ke akhir =====
            int last = blokStart.get(blokStart.size() - 1);

            jadwal[last][kelas] = guruFirst;
            jadwal[last + 1][kelas] = guruFirst;

            return; //
        }
    }

    static void swap2WithSingle(String[][] jadwal, int[][] hariRange) {

        Random rand = new Random();
        int iterasi = 0;
        int maxPercobaan = 1;
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

            iterasi++;
        }

//        if (percobaan >= maxPercobaan) {
//            System.out.println("[swap2WithSingle] gagal setelah " + maxPercobaan + " percobaan");
//        }
    }

    static void swap3WithDoubleAndSingle(String[][] jadwal, int[][] hariRange) {

        Random rand = new Random();
        int iterasi = 0;
        int maxPercobaan = 1;
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


    static String[][] hillClimbingTabuHardConstrain(
            String[][] jadwal,
            int[][] hariRange,
            Set<String> guruPJOK) {

        Random rand = new Random();
        Queue<Integer> tabuList = new LinkedList<>();
        int tabuSize = 3;

        String[][] current = copyJadwal(jadwal);

        int penaltiSekarang = hitungTotalSemuaPenaltiHardConstrain(current, hariRange, guruPJOK);
        int penaltiPJOKSekarang = hitungPenaltiPJOK(current, hariRange, guruPJOK);
        int penaltiGuru43Sekarang = hitungPenaltiGuruID43(current, hariRange);
        int penaltiMaxGuruPerHariSekarang = hitungPenaltiMaxGuruPerHari(current, hariRange, 6);

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
//                System.out.println("Target tercapai");
                System.out.println("Selesai di iterasi global: " + iterasiGlobal[0] + " | Penalti akhir: " + penaltiSekarang);
                break;
            }

            String[][] neighbor = copyJadwal(current);

            //tabu
            int move;

            do {
                move = rand.nextInt(6);
            } while (tabuList.contains(move));

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
            } else {
                swap3Random(neighbor, hariRange);
            }

            tabuList.offer(move);
            if (tabuList.size() > tabuSize) {
                tabuList.poll();
            }

            int penaltiBaru = hitungTotalSemuaPenaltiHardConstrain(neighbor, hariRange, guruPJOK);

            int penaltiPJOKBaru = hitungPenaltiPJOK(neighbor, hariRange, guruPJOK);
            int penaltiGuru43Baru = hitungPenaltiGuruID43(neighbor, hariRange);
            int penaltiMaxGuruPerHariBaru = hitungPenaltiMaxGuruPerHari(neighbor, hariRange, 6);

//            System.out.println(
//                    "Total: " + penaltiSekarang + " -> " + penaltiBaru +
//                            " | PJOK: " + penaltiPJOKSekarang + " -> " + penaltiPJOKBaru
//            );

            if (penaltiBaru <= penaltiSekarang
                    && penaltiPJOKSekarang >= penaltiPJOKBaru
                    && penaltiGuru43Sekarang >= penaltiGuru43Baru
                    && penaltiMaxGuruPerHariSekarang >= penaltiMaxGuruPerHariBaru){

                for (int i = 0; i < current.length; i++) {
                    for (int j = 0; j < current[i].length; j++) {
                        current[i][j] = neighbor[i][j];
                    }
                }
                penaltiSekarang = penaltiBaru;
                penaltiPJOKSekarang = penaltiPJOKBaru;
                penaltiGuru43Sekarang = penaltiGuru43Baru;
                penaltiMaxGuruPerHariSekarang = penaltiMaxGuruPerHariBaru;

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

    static String[][] hillClimbingRandomMGMPKamis(
            String[][] jadwal,
            int[][] hariRange,
            Set<String> guruPJOK,
            Set<String> MGMPkamis) {

        Random rand = new Random();

        String[][] current = copyJadwal(jadwal);

        int penaltiSekarang = hitungPenaltiMGMPKamis(jadwal, hariRange, MGMPkamis);
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
//                System.out.println("Target tercapai");
                System.out.println("Selesai di iterasi global: " + iterasiGlobal[0] + " | Penalti akhir: " + penaltiSekarang);
                break;
            }

            String[][] neighbor = copyJadwal(current);

            // random
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


            int penaltiBaru = hitungPenaltiMGMPKamis(neighbor, hariRange, MGMPkamis);
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

    static String[][] hillClimbingRandomHardConstrain(
            String[][] jadwal,
            int[][] hariRange,
            Set<String> guruPJOK,
            Set<String> MGMPkamis) {

        Random rand = new Random();

        String[][] current = copyJadwal(jadwal);

        int penaltiSekarang = hitungTotalSemuaPenaltiHardConstrain(current, hariRange, guruPJOK);
        int penaltiPJOKSekarang = hitungPenaltiPJOK(current, hariRange, guruPJOK);
        int penaltiMGMPKamisSekarang = hitungPenaltiMGMPKamis(jadwal, hariRange, MGMPkamis);


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
//                System.out.println("Target tercapai");
                System.out.println("Selesai di iterasi global: " + iterasiGlobal[0] + " | Penalti akhir: " + penaltiSekarang);
                break;
            }

            String[][] neighbor = copyJadwal(current);

            // random
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


            int penaltiBaru = hitungTotalSemuaPenaltiHardConstrain(neighbor, hariRange, guruPJOK);
            int penaltiPJOKBaru = hitungPenaltiPJOK(neighbor, hariRange, guruPJOK);
            int penaltiMGMPKamisBaru = hitungPenaltiMGMPKamis(neighbor, hariRange, MGMPkamis);

//            System.out.println(
//                    "Total: " + penaltiSekarang + " -> " + penaltiBaru +
//                            " | PJOK: " + penaltiPJOKSekarang + " -> " + penaltiPJOKBaru
//            );

            if (penaltiBaru <= penaltiSekarang
                    && penaltiPJOKSekarang >= penaltiPJOKBaru
                    && penaltiMGMPKamisSekarang >= penaltiMGMPKamisBaru){

                for (int i = 0; i < current.length; i++) {
                    for (int j = 0; j < current[i].length; j++) {
                        current[i][j] = neighbor[i][j];
                    }
                }
                penaltiSekarang = penaltiBaru;
                penaltiPJOKSekarang = penaltiPJOKBaru;
                penaltiMGMPKamisSekarang = penaltiMGMPKamisBaru;
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

//        Queue<Integer> tabuList = new LinkedList<>();
//        int tabuSize = 2;

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

            //tabu

//            int move;
//            do {
//                move = rand.nextInt(5);
//            } while (tabuList.contains(move));
//
//            if (move == 0) {
//                swap1Random(neighbor, hariRange);
//            } else if (move == 1) {
//                swap2Random(neighbor, hariRange);
//            } else if (move == 2) {
//                swap2WithSingle(neighbor, hariRange);
//            }else if (move == 3) {
//                 swap2RandomPJOK(neighbor, hariRange, guruPJOK);
//            } else {
//                rotationRandom(neighbor, hariRange);
//            }
//            tabuList.offer(move);
//            if (tabuList.size() > tabuSize) {
//                tabuList.poll();
//            }

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

    static String[][] LAHCRandom(
            String[][] jadwal,
            int[][] hariRange,
            Set<String> guruPJOK,
            Set<String> MGMPsenin,
            Set<String> MGMPselasa,
            Set<String> MGMPrabu,
            Set<String> MGMPkamis,
            Set<String> guruMatematika) {


        String[][] current = copyJadwal(jadwal);
        Random rand = new Random();


        int penaltiAwal = hitungTotalSemuaPenalti(
                current, hariRange,
                guruPJOK,
                MGMPsenin, MGMPselasa, MGMPrabu, MGMPkamis,
                guruMatematika
        );


        int l = 100;
        int[] fitnessHistory = new int[l];

        // Mengisi slot isi
        for (int i = 0; i < l; i++) {
            fitnessHistory[i] = penaltiAwal;
        }

        int iter = 0;

        System.out.println("Penalti awal: " + penaltiAwal);

        while (penaltiAwal > 0) {
            String[][] neighbor = copyJadwal(current);

            int move;

           //LLH



            int kandidatSolusi = hitungTotalSemuaPenalti(
                    neighbor, hariRange,
                    guruPJOK,
                    MGMPsenin, MGMPselasa, MGMPrabu, MGMPkamis,
                    guruMatematika
            );




            System.out.println(
                    "Iter " + iter +
                            " | Current: " + penaltiAwal +
                            " | New: " + kandidatSolusi +
                            " | History[" + iter% fitnessHistory.length + "]: " + fitnessHistory[iter % fitnessHistory.length]
            );
            if (kandidatSolusi <= fitnessHistory[iter % fitnessHistory.length] || kandidatSolusi <= penaltiAwal) {
                for (int i = 0; i < current.length; i++) {
                    for (int j = 0; j < current[i].length; j++) {
                        current[i][j] = neighbor[i][j];
                    }
                }
                penaltiAwal = kandidatSolusi;
                fitnessHistory[iter % fitnessHistory.length] = kandidatSolusi;
//                System.out.println("Diterima");
            } else {
                for (int i = 0; i < neighbor.length; i++) {
                    for (int j = 0; j < neighbor[i].length; j++) {
                        neighbor[i][j] = current[i][j];
                    }
                }
                System.out.println("Ditolak");
            }


            iter++;


            if (penaltiAwal <= 0) {
                System.out.println("Solusi optimal");
                break;
            }


        }
        System.out.println("Selesai di iterasi: " + iter);
        System.out.println("Penalti akhir: " + penaltiAwal);


        return current;
    }

}



