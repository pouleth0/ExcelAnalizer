package org.studproject.view;

import org.studproject.controller.LoadexcelMod01;

import javax.swing.*;

public class StartApp {
    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> new LoadexcelMod01());
    }
}
