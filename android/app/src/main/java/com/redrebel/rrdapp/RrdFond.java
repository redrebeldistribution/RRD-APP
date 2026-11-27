package com.redrebel.rrdapp;

import android.content.Context;
import android.content.Intent;
import android.content.SharedPreferences;
import android.net.Uri;
import android.os.Build;
import android.os.PowerManager;
import android.provider.Settings;

import com.getcapacitor.JSArray;
import com.getcapacitor.JSObject;
import com.getcapacitor.Plugin;
import com.getcapacitor.PluginCall;
import com.getcapacitor.PluginMethod;
import com.getcapacitor.annotation.CapacitorPlugin;

import java.io.File;

/** Pont entre les écrans de l'application et l'envoi en arrière-plan (EnvoiFond). */
@CapacitorPlugin(name = "RrdFond")
public class RrdFond extends Plugin {
    static volatile boolean premierPlan = false;

    @Override public void load() { premierPlan = true; EnvoiFond.planifierHoraire(getContext()); }
    @Override protected void handleOnResume() { premierPlan = true; }
    @Override protected void handleOnPause() { premierPlan = false; EnvoiFond.planifier(getContext(), 45); }
    @Override protected void handleOnDestroy() { premierPlan = false; }

    private File dir() { return EnvoiFond.dossier(getContext()); }

    /** Adresse du bureau et jeton du téléphone : lus par la tâche au moment de l'envoi. */
    @PluginMethod
    public void regler(PluginCall call) {
        EnvoiFond.prefs(getContext()).edit().putString("url", call.getString("url", "")).putString("jeton", call.getString("jeton", "")).apply();
        call.resolve();
    }

    @PluginMethod
    public void ajouter(PluginCall call) {
        String id = call.getString("id"), corps = call.getString("corps");
        long ordre = call.getData().optLong("ordre", 0);
        if (!EnvoiCoeur.idValide(id) || corps == null || corps.indexOf(EnvoiCoeur.JETON) < 0) { call.reject("élément invalide"); return; }
        try {
            EnvoiCoeur.retirerAttente(dir(), id);
            EnvoiCoeur.ecrire(EnvoiCoeur.fichierAttente(dir(), ordre, id), corps);
            call.resolve();
        } catch (Exception e) { call.reject("écriture impossible : " + e.getMessage()); }
    }

    /** L'élément est réglé dans l'application : on oublie sa copie et la réponse du bureau. */
    @PluginMethod
    public void retirer(PluginCall call) {
        String id = call.getString("id");
        if (!EnvoiCoeur.idValide(id)) { call.reject("id invalide"); return; }
        EnvoiCoeur.retirerAttente(dir(), id);
        EnvoiCoeur.fichierReponse(dir(), id).delete();
        call.resolve();
    }

    @PluginMethod
    public void resultat(PluginCall call) {
        String id = call.getString("id");
        JSObject o = new JSObject();
        File f = EnvoiCoeur.idValide(id) ? EnvoiCoeur.fichierReponse(dir(), id) : null;
        try { o.put("rep", f != null && f.isFile() ? EnvoiCoeur.lire(f) : ""); } catch (Exception e) { o.put("rep", ""); }
        call.resolve(o);
    }

    @PluginMethod
    public void etat(PluginCall call) {
        JSArray att = new JSArray(), faits = new JSArray();
        long octets = 0;
        for (File f : EnvoiCoeur.liste(dir(), "a_")) { String id = EnvoiCoeur.idDe(f.getName()); if (id != null) { att.put(id); octets += f.length(); } }
        for (File f : EnvoiCoeur.liste(dir(), "r_")) { String id = EnvoiCoeur.idDe(f.getName()); if (id != null) faits.put(id); }
        SharedPreferences p = EnvoiFond.prefs(getContext());
        JSObject o = new JSObject();
        o.put("attente", att); o.put("faits", faits); o.put("octets", octets);
        o.put("dernier_ms", p.getLong("dernier_ms", 0)); o.put("dernier_code", p.getInt("dernier_code", 0));
        o.put("dernier_envoyes", p.getInt("dernier_envoyes", 0)); o.put("dernier_refuses", p.getInt("dernier_refuses", 0));
        o.put("dernier_detail", p.getString("dernier_detail", "")); o.put("total_envoyes", p.getInt("total_envoyes", 0));
        o.put("batterie_libre", batterieLibre());
        call.resolve(o);
    }

    @PluginMethod
    public void vider(PluginCall call) {
        File[] fs = dir().listFiles();
        if (fs != null) for (File f : fs) f.delete();
        call.resolve();
    }

    private boolean batterieLibre() {
        try {
            if (Build.VERSION.SDK_INT < 23) return true;
            PowerManager pm = (PowerManager) getContext().getSystemService(Context.POWER_SERVICE);
            return pm != null && pm.isIgnoringBatteryOptimizations(getContext().getPackageName());
        } catch (Exception e) { return false; }
    }

    /** Ouvre la fiche de l'application dans les paramètres du téléphone (réglage de la batterie). */
    @PluginMethod
    public void reglages(PluginCall call) {
        try {
            Intent i = new Intent(Settings.ACTION_APPLICATION_DETAILS_SETTINGS, Uri.parse("package:" + getContext().getPackageName()));
            i.addFlags(Intent.FLAG_ACTIVITY_NEW_TASK);
            getContext().startActivity(i);
            call.resolve();
        } catch (Exception e) { call.reject("paramètres indisponibles"); }
    }
}
