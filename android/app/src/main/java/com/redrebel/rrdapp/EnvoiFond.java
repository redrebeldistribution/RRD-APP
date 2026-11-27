package com.redrebel.rrdapp;

import android.content.Context;
import android.content.SharedPreferences;

import androidx.annotation.NonNull;
import androidx.work.BackoffPolicy;
import androidx.work.Constraints;
import androidx.work.ExistingPeriodicWorkPolicy;
import androidx.work.ExistingWorkPolicy;
import androidx.work.NetworkType;
import androidx.work.OneTimeWorkRequest;
import androidx.work.PeriodicWorkRequest;
import androidx.work.WorkManager;
import androidx.work.Worker;
import androidx.work.WorkerParameters;

import java.io.File;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.locks.ReentrantLock;

/**
 * Tâche d'Android qui envoie la file d'attente quand il y a du réseau, même si l'application est fermée.
 * Tant que l'application est à l'écran, elle ne fait rien : l'application envoie elle-même.
 */
public class EnvoiFond extends Worker {
    static final String PREFS = "rrd_fond", UNIQUE = "rrd-envoi", HORAIRE = "rrd-envoi-h";
    private static final ReentrantLock VERROU = new ReentrantLock();

    public EnvoiFond(@NonNull Context c, @NonNull WorkerParameters p) { super(c, p); }

    static File dossier(Context c) { File d = new File(c.getFilesDir(), "rrd_fond"); if (!d.isDirectory()) d.mkdirs(); return d; }
    static SharedPreferences prefs(Context c) { return c.getSharedPreferences(PREFS, Context.MODE_PRIVATE); }

    private static Constraints reseau() { return new Constraints.Builder().setRequiredNetworkType(NetworkType.CONNECTED).build(); }

    /** À appeler quand l'application quitte l'écran : un passage dès qu'il y a du réseau (après un court délai). */
    static void planifier(Context c, long delaiSecondes) {
        try {
            if (EnvoiCoeur.liste(dossier(c), "a_").isEmpty()) return;
            OneTimeWorkRequest r = new OneTimeWorkRequest.Builder(EnvoiFond.class).setConstraints(reseau())
                .setInitialDelay(delaiSecondes, TimeUnit.SECONDS).setBackoffCriteria(BackoffPolicy.LINEAR, 2, TimeUnit.MINUTES).build();
            WorkManager.getInstance(c).enqueueUniqueWork(UNIQUE, ExistingWorkPolicy.REPLACE, r);
        } catch (Exception e) { /* jamais de panne de l'application pour cela */ }
    }
    /** Filet de sécurité : un passage par heure quand il y a du réseau (redémarrage du téléphone, tâche perdue). */
    static void planifierHoraire(Context c) {
        try {
            PeriodicWorkRequest r = new PeriodicWorkRequest.Builder(EnvoiFond.class, 1, TimeUnit.HOURS).setConstraints(reseau()).build();
            WorkManager.getInstance(c).enqueueUniquePeriodicWork(HORAIRE, ExistingPeriodicWorkPolicy.KEEP, r);
        } catch (Exception e) { }
    }

    @NonNull @Override
    public Result doWork() {
        if (RrdFond.premierPlan) return Result.success();
        if (!VERROU.tryLock()) return Result.success();   // un autre passage est en cours
        try {
            Context c = getApplicationContext();
            SharedPreferences p = prefs(c);
            EnvoiCoeur.Bilan b = EnvoiCoeur.vider(dossier(c), p.getString("url", ""), p.getString("jeton", ""), () -> isStopped() || RrdFond.premierPlan);
            p.edit().putLong("dernier_ms", System.currentTimeMillis()).putInt("dernier_code", b.code).putInt("dernier_envoyes", b.envoyes)
                .putInt("dernier_refuses", b.refuses).putString("dernier_detail", b.detail == null ? "" : b.detail)
                .putInt("total_envoyes", p.getInt("total_envoyes", 0) + b.envoyes).apply();
            boolean reste = !EnvoiCoeur.liste(dossier(c), "a_").isEmpty();
            if (b.code == EnvoiCoeur.REESSAYER || (reste && isStopped() && !RrdFond.premierPlan)) return Result.retry();
            return Result.success();
        } catch (Exception e) {
            return Result.retry();
        } finally { VERROU.unlock(); }
    }
}
