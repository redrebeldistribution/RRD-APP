package com.redrebel.rrdapp;

import android.os.Bundle;

import com.getcapacitor.BridgeActivity;

public class MainActivity extends BridgeActivity {
    @Override
    public void onCreate(Bundle savedInstanceState) {
        registerPlugin(RrdFond.class);   // envoi de la file d'attente quand l'application est fermée
        super.onCreate(savedInstanceState);
    }
}
