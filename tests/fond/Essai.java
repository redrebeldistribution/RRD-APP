import com.redrebel.rrdapp.EnvoiCoeur;
import java.io.File;
import java.nio.file.Files;
import java.util.*;

public class Essai {
    static int n = 0, bad = 0;
    static void v(String nom, boolean ok) { n++; if (!ok) { bad++; System.out.println("ÉCHEC  " + nom); } else System.out.println("ok     " + nom); }
    static File neuf() throws Exception { File d = Files.createTempDirectory("fond").toFile(); return d; }
    static void att(File d, long ordre, String id, String extra) throws Exception { EnvoiCoeur.ecrire(EnvoiCoeur.fichierAttente(d, ordre, id), "{\"action\":\"photo\",\"jeton\":\"@@JETON@@\"," + extra + "}"); }
    static String[] noms(File d) { String[] s = d.list(); Arrays.sort(s); return s; }
    public static void main(String[] a) throws Exception {
        String url = a[0];
        // analyser
        v("ok simple", Arrays.equals(EnvoiCoeur.analyser("{\"ok\":true,\"data\":{\"recu\":true,\"cid\":\"P1\"}}"), new String[]{"1",""}));
        v("refus avec code", Arrays.equals(EnvoiCoeur.analyser("{\"ok\":false,\"erreur\":\"Photo trop ancienne (plus de 31 jours).\",\"code\":\"\"}"), new String[]{"0",""}));
        v("refus VERIF + controles", Arrays.equals(EnvoiCoeur.analyser("{\"ok\":false,\"erreur\":\"x \\\"code\\\":\\\"FAUX\\\" y\",\"code\":\"VERIF\",\"controles\":[{\"k\":\"a\",\"ok\":true},[1,2,{\"z\":null}]]}"), new String[]{"0","VERIF"}));
        v("ok imbriqué ne compte pas", Arrays.equals(EnvoiCoeur.analyser(" {\"data\":{\"ok\":false,\"code\":\"X\"},\"ok\":true} \n"), new String[]{"1",""}));
        v("html", EnvoiCoeur.analyser("<!DOCTYPE html><html>Erreur</html>") == null);
        v("vide", EnvoiCoeur.analyser("") == null && EnvoiCoeur.analyser("{}") == null && EnvoiCoeur.analyser("{\"ok\":true") == null);
        v("nombres et unicode", Arrays.equals(EnvoiCoeur.analyser("{\"ok\":true,\"data\":{\"n\":-1.5e3,\"t\":\"caf\\u00e9 \\n\",\"v\":[],\"o\":{}}}"), new String[]{"1",""}));
        v("ids", EnvoiCoeur.idValide("ESA-20261127-ABCD") && !EnvoiCoeur.idValide("a_b") && !EnvoiCoeur.idValide("../x") && "P1".equals(EnvoiCoeur.idDe("a_0000000000005_P1.json")) && "ESA-1-A".equals(EnvoiCoeur.idDe("r_ESA-1-A.json")));

        // 1) trois éléments, ordre respecté, jeton mis à la place
        File d = neuf();
        att(d, 300, "P3", "\"payload\":{\"cid\":\"P3\",\"photo\":\"data:image/jpeg;base64,AAAA/9j+==\"}");
        att(d, 100, "P1", "\"payload\":{\"cid\":\"P1\",\"note\":\"é \\\"@@JETON@@\\\" ne doit pas changer\"}");
        att(d, 200, "ESA-20261127-ABCD", "\"payload\":{\"cid\":\"R2\"}");
        EnvoiCoeur.Bilan b = EnvoiCoeur.vider(d, url + "?mode=ok", "JET\"ON\\1", null);
        v("tout envoyé", b.code == EnvoiCoeur.FINI && b.envoyes == 3 && b.refuses == 0);
        v("copies retirées, réponses écrites", Arrays.equals(noms(d), new String[]{"r_ESA-20261127-ABCD.json","r_P1.json","r_P3.json"}));
        String r1 = EnvoiCoeur.lire(EnvoiCoeur.fichierReponse(d, "P1"));
        v("réponse gardée telle quelle", r1.contains("\"ok\":true") && r1.contains("\"cid\":\"P1\""));
        v("ordre 1,2,3 + jeton + note intacte", r1.contains("\"rang\":1") && EnvoiCoeur.lire(EnvoiCoeur.fichierReponse(d, "ESA-20261127-ABCD")).contains("\"rang\":2") && EnvoiCoeur.lire(EnvoiCoeur.fichierReponse(d, "P3")).contains("\"rang\":3")
            && r1.contains("\"jeton_ok\":true") && r1.contains("\"note_ok\":true"));
        // 2) refus définitif au milieu : on continue
        d = neuf(); att(d, 1, "A", "\"payload\":{\"cid\":\"A\"}"); att(d, 2, "B", "\"payload\":{\"cid\":\"B\",\"refuse\":true}"); att(d, 3, "C", "\"payload\":{\"cid\":\"C\"}");
        b = EnvoiCoeur.vider(d, url + "?mode=ok", "J", null);
        v("refus définitif : file continue", b.code == EnvoiCoeur.FINI && b.envoyes == 2 && b.refuses == 1 && EnvoiCoeur.lire(EnvoiCoeur.fichierReponse(d, "B")).contains("trop ancienne") && noms(d).length == 3);
        // 3) serveur occupé : on s'arrête, rien n'est perdu
        d = neuf(); att(d, 1, "A", "\"payload\":{\"cid\":\"A\"}"); att(d, 2, "B", "\"payload\":{\"cid\":\"B\",\"occupe\":true}"); att(d, 3, "C", "\"payload\":{\"cid\":\"C\"}");
        b = EnvoiCoeur.vider(d, url + "?mode=ok", "J", null);
        v("serveur occupé : arrêt, B et C gardés", b.code == EnvoiCoeur.REESSAYER && b.envoyes == 1 && Arrays.equals(noms(d), new String[]{"a_0000000000002_B.json","a_0000000000003_C.json","r_A.json"}));
        // 4) code refusé
        d = neuf(); att(d, 1, "A", "\"payload\":{\"cid\":\"A\"}");
        b = EnvoiCoeur.vider(d, url + "?mode=ok", "MAUVAIS", null);
        v("jeton refusé : arrêt sans perte", b.code == EnvoiCoeur.ARRET && noms(d).length == 1 && noms(d)[0].startsWith("a_"));
        // 5) page d'erreur html, puis panne de réseau
        b = EnvoiCoeur.vider(d, url + "?mode=html", "J", null);
        v("réponse illisible : on réessaiera", b.code == EnvoiCoeur.REESSAYER && noms(d).length == 1);
        b = EnvoiCoeur.vider(d, url + "?mode=500", "J", null);
        v("erreur 500 : on réessaiera", b.code == EnvoiCoeur.REESSAYER && noms(d).length == 1);
        b = EnvoiCoeur.vider(d, "http://127.0.0.1:9/exec", "J", null);
        v("pas de réseau : on réessaiera", b.code == EnvoiCoeur.REESSAYER && noms(d).length == 1);
        // 6) interruption (application revenue à l'écran)
        final int[] k = {0};
        d = neuf(); att(d, 1, "A", "\"payload\":{\"cid\":\"A\"}"); att(d, 2, "B", "\"payload\":{\"cid\":\"B\"}");
        b = EnvoiCoeur.vider(d, url + "?mode=ok", "J", () -> k[0]++ >= 1);
        v("interrompu après le premier", b.code == EnvoiCoeur.INTERROMPU && b.envoyes == 1 && noms(d).length == 2);
        // 7) fichier abîmé, pas de jeton, dossier vide
        d = neuf(); EnvoiCoeur.ecrire(new File(d, "a_0000000000001_X.json"), "{pas de marque}"); att(d, 2, "B", "\"payload\":{\"cid\":\"B\"}");
        b = EnvoiCoeur.vider(d, url + "?mode=ok", "J", null);
        v("fichier abîmé écarté, le reste part", b.code == EnvoiCoeur.FINI && b.envoyes == 1 && Arrays.equals(noms(d), new String[]{"r_B.json"}));
        v("sans jeton : rien", EnvoiCoeur.vider(d, url, "", null).code == EnvoiCoeur.FINI);
        v("dossier vide", EnvoiCoeur.vider(neuf(), url + "?mode=ok", "J", null).envoyes == 0);
        // 8) gros élément (photo de 600 ko) et remplacement d'une copie
        d = neuf(); StringBuilder big = new StringBuilder(); for (int i = 0; i < 600000; i++) big.append((char) ('A' + i % 26));
        att(d, 5, "G", "\"payload\":{\"cid\":\"G\",\"photo\":\"data:image/jpeg;base64," + big + "\"}");
        b = EnvoiCoeur.vider(d, url + "?mode=ok", "J", null);
        v("gros élément", b.envoyes == 1 && EnvoiCoeur.lire(EnvoiCoeur.fichierReponse(d, "G")).contains("\"taille\":6000"));
        att(d, 1, "Z", "\"payload\":{}"); EnvoiCoeur.retirerAttente(d, "Z"); v("retirerAttente", noms(d).length == 1);
        System.out.println(n - bad + "/" + n + " réussis");
        System.exit(bad == 0 ? 0 : 1);
    }
}
