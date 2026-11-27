package com.redrebel.rrdapp;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * Envoi de la file d'attente sans l'écran de l'application (application fermée).
 *
 * Le dossier contient, pour chaque élément en attente (rapport, photo, bordereau, opération de camion) :
 *   a_<ordre sur 13 chiffres>_<id>.json   la requête exacte que l'application aurait envoyée, avec @@JETON@@ à la place du jeton ;
 *   r_<id>.json                            la réponse du bureau, quand elle est définitive (reçu ou refusé).
 * L'application lit les réponses à sa prochaine ouverture et met ses écrans à jour ; elle seule retire les r_.
 * Le bureau ne garde jamais deux fois le même élément : un double envoi est sans conséquence.
 *
 * Cette classe n'utilise rien d'Android : elle se teste sur un ordinateur.
 */
public final class EnvoiCoeur {
    public static final int FINI = 0, REESSAYER = 1, ARRET = 2, INTERROMPU = 3;
    public static final String JETON = "\"@@JETON@@\"";

    public interface Stop { boolean arret(); }

    public static final class Bilan {
        public int code = FINI, envoyes = 0, refuses = 0;
        public String detail = "";
    }

    private EnvoiCoeur() {}

    public static boolean idValide(String id) { return id != null && id.matches("[A-Za-z0-9-]{1,80}"); }
    public static File fichierAttente(File dir, long ordre, String id) { return new File(dir, String.format(java.util.Locale.US, "a_%013d_%s.json", Math.max(0, ordre), id)); }
    public static File fichierReponse(File dir, String id) { return new File(dir, "r_" + id + ".json"); }
    public static String idDe(String nom) {   // a_0000000000123_ID.json ou r_ID.json
        if (!nom.endsWith(".json")) return null;
        String s = nom.substring(0, nom.length() - 5);
        if (s.startsWith("a_")) { int i = s.indexOf('_', 2); return i < 0 ? null : s.substring(i + 1); }
        if (s.startsWith("r_")) return s.substring(2);
        return null;
    }
    public static List<File> liste(File dir, String prefixe) {
        List<File> l = new ArrayList<>();
        File[] fs = dir.listFiles();
        if (fs != null) for (File f : fs) if (f.isFile() && f.getName().startsWith(prefixe) && f.getName().endsWith(".json")) l.add(f);
        File[] t = l.toArray(new File[0]);
        Arrays.sort(t, (a, b) -> a.getName().compareTo(b.getName()));
        return new ArrayList<>(Arrays.asList(t));
    }
    public static void retirerAttente(File dir, String id) { for (File f : liste(dir, "a_")) if (id.equals(idDe(f.getName()))) f.delete(); }

    public static String lire(File f) throws IOException {
        try (InputStream in = new FileInputStream(f)) { return new String(tout(in), StandardCharsets.UTF_8); }
    }
    /** Écriture complète ou rien : fichier provisoire puis changement de nom. */
    public static void ecrire(File f, String texte) throws IOException {
        File tmp = new File(f.getParentFile(), "t_" + f.getName());
        try (OutputStream out = new FileOutputStream(tmp)) { out.write(texte.getBytes(StandardCharsets.UTF_8)); out.flush(); }
        if (f.exists()) f.delete();
        if (!tmp.renameTo(f)) { tmp.delete(); throw new IOException("écriture impossible : " + f.getName()); }
    }
    private static byte[] tout(InputStream in) throws IOException {
        ByteArrayOutputStream o = new ByteArrayOutputStream(); byte[] b = new byte[16384]; int n;
        while ((n = in.read(b)) > 0) o.write(b, 0, n);
        return o.toByteArray();
    }
    static String citer(String s) {   // chaîne JSON
        StringBuilder b = new StringBuilder("\"");
        for (int i = 0; i < s.length(); i++) {
            char c = s.charAt(i);
            if (c == '"' || c == '\\') b.append('\\').append(c);
            else if (c < 0x20) b.append(String.format("\\u%04x", (int) c));
            else b.append(c);
        }
        return b.append('"').toString();
    }

    /** POST comme l'application ; le serveur (Apps Script) répond par une redirection qu'on suit en GET. */
    public static String poster(String url, String corps) throws IOException {
        byte[] data = corps.getBytes(StandardCharsets.UTF_8);
        String u = url; boolean post = true;
        for (int saut = 0; saut < 6; saut++) {
            HttpURLConnection c = (HttpURLConnection) new URL(u).openConnection();
            try {
                c.setInstanceFollowRedirects(false);
                c.setConnectTimeout(30000); c.setReadTimeout(120000);
                c.setRequestProperty("Accept", "application/json, text/plain, */*");
                if (post) {
                    c.setRequestMethod("POST"); c.setDoOutput(true);
                    c.setRequestProperty("Content-Type", "text/plain;charset=utf-8");
                    c.setFixedLengthStreamingMode(data.length);
                    try (OutputStream o = c.getOutputStream()) { o.write(data); }
                } else c.setRequestMethod("GET");
                int st = c.getResponseCode();
                if (st == 301 || st == 302 || st == 303 || st == 307 || st == 308) {
                    String loc = c.getHeaderField("Location");
                    if (loc == null || loc.isEmpty()) throw new IOException("redirection sans adresse");
                    u = new URL(new URL(u), loc).toString();
                    if (st != 307 && st != 308) post = false;
                    continue;
                }
                if (st < 200 || st >= 300) throw new IOException("réponse " + st);
                try (InputStream in = c.getInputStream()) { return new String(tout(in), StandardCharsets.UTF_8); }
            } finally { c.disconnect(); }
        }
        throw new IOException("trop de redirections");
    }

    /** Lit seulement ce qu'il faut de la réponse {"ok":true,…} ou {"ok":false,"erreur":…,"code":…} : [ok, code] ; null si ce n'est pas une réponse du bureau. */
    public static String[] analyser(String rep) {
        try {
            Lecteur l = new Lecteur(rep);
            l.blancs(); l.attendre('{');
            Boolean ok = null; String code = "";
            l.blancs();
            if (l.voir() != '}') {
                for (;;) {
                    l.blancs(); String cle = l.chaine(); l.blancs(); l.attendre(':'); l.blancs();
                    if (cle.equals("ok")) { String v = l.mot(); if (v.equals("true")) ok = Boolean.TRUE; else if (v.equals("false")) ok = Boolean.FALSE; else return null; }
                    else if (cle.equals("code") && l.voir() == '"') code = l.chaine();
                    else l.sauter();
                    l.blancs();
                    if (l.voir() == ',') { l.i++; continue; }
                    l.attendre('}'); break;
                }
            } else l.i++;
            l.blancs();
            if (l.i != rep.length() || ok == null) return null;
            return new String[] { ok ? "1" : "0", code };
        } catch (RuntimeException e) { return null; }
    }
    private static final class Lecteur {
        final String s; int i = 0;
        Lecteur(String s) { this.s = s; }
        char voir() { if (i >= s.length()) throw new IllegalStateException(); return s.charAt(i); }
        void blancs() { while (i < s.length() && " \t\r\n".indexOf(s.charAt(i)) >= 0) i++; }
        void attendre(char c) { if (voir() != c) throw new IllegalStateException(); i++; }
        String chaine() {
            attendre('"'); StringBuilder b = new StringBuilder();
            for (;;) {
                char c = voir(); i++;
                if (c == '"') return b.toString();
                if (c == '\\') {
                    char e = voir(); i++;
                    switch (e) {
                        case 'n': b.append('\n'); break; case 't': b.append('\t'); break; case 'r': b.append('\r'); break;
                        case 'b': b.append('\b'); break; case 'f': b.append('\f'); break;
                        case 'u': b.append((char) Integer.parseInt(s.substring(i, i + 4), 16)); i += 4; break;
                        default: b.append(e);
                    }
                } else b.append(c);
            }
        }
        String mot() {   // true, false, null ou un nombre
            int d = i;
            while (i < s.length()) { char c = s.charAt(i); if (Character.isLetterOrDigit(c) || c == '+' || c == '-' || c == '.') i++; else break; }
            if (i == d) throw new IllegalStateException();
            return s.substring(d, i);
        }
        void sauter() {   // une valeur quelconque
            char c = voir();
            if (c == '"') { chaine(); return; }
            if (c == '{' || c == '[') {
                char fin = c == '{' ? '}' : ']'; i++; blancs();
                if (voir() == fin) { i++; return; }
                for (;;) {
                    blancs();
                    if (c == '{') { chaine(); blancs(); attendre(':'); blancs(); }
                    sauter(); blancs();
                    if (voir() == ',') { i++; continue; }
                    attendre(fin); return;
                }
            }
            mot();
        }
    }

    /** Envoie tout ce qui attend, du plus ancien au plus récent. S'arrête au premier échec de réseau ou de serveur. */
    public static Bilan vider(File dir, String url, String jeton, Stop stop) {
        Bilan b = new Bilan();
        if (url == null || url.isEmpty() || jeton == null || jeton.isEmpty()) { b.detail = "pas de code actif"; return b; }
        for (int tour = 0; tour < 1000; tour++) {
            List<File> fs = liste(dir, "a_");
            if (fs.isEmpty()) return b;
            for (File f : fs) {
                if (stop != null && stop.arret()) { b.code = INTERROMPU; return b; }
                if (!f.exists()) continue;   // l'application l'a envoyé elle-même entre-temps
                String id = idDe(f.getName()), corps;
                try { corps = lire(f); } catch (IOException e) { continue; }
                int p = corps.indexOf(JETON);
                if (id == null || !idValide(id) || p < 0) { f.delete(); continue; }   // fichier abîmé : l'application le récrira
                corps = corps.substring(0, p) + citer(jeton) + corps.substring(p + JETON.length());
                String rep;
                try { rep = poster(url, corps); }
                catch (IOException e) { b.code = REESSAYER; b.detail = "réseau : " + e.getMessage(); return b; }
                String[] a = analyser(rep);
                if (a == null) { b.code = REESSAYER; b.detail = "réponse illisible du serveur"; return b; }
                if (a[0].equals("0")) {
                    if (a[1].equals("SERVEUR") || a[1].equals("CONFIG")) { b.code = REESSAYER; b.detail = "serveur occupé"; return b; }
                    if (a[1].equals("NON_ENREGISTRE") || a[1].equals("DESACTIVE")) { b.code = ARRET; b.detail = "code du téléphone refusé"; return b; }
                }
                try { ecrire(fichierReponse(dir, id), rep); }
                catch (IOException e) { b.code = REESSAYER; b.detail = "mémoire du téléphone : " + e.getMessage(); return b; }
                f.delete();
                if (a[0].equals("1")) b.envoyes++; else b.refuses++;
            }
        }
        return b;
    }
}
