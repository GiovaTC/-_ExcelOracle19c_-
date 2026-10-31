package com.exceloracle;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.awt.Font;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.sql.*;
import java.text.SimpleDateFormat;

public class Main extends JFrame
{
    private static final String URL = "jdbc:oracle:thin:@//localhost:1521/orcl";
    private static final String USUARIO = "system";
    private static final String CLAVE = "Tapiero123";

    private JTable tabla;

    private DefaultTableModel modelo;

    public Main()
    {
        setTitle("Excel - Oracle 19c");
        setSize(900, 550);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setLocationRelativeTo(null);
        inicializarInterfaz();
    }

    // INTERFAZ GRAFICA.
    private void inicializarInterfaz() {

        JPanel panelPrincipal = new JPanel(new BorderLayout(10, 10));

        panelPrincipal.setBorder(BorderFactory.createEmptyBorder(15, 15, 15, 15));

        JLabel titulo = new JLabel("GENERADOR DE ARCHIVOS | EXCEL | ORACLE 19C", SwingConstants.CENTER);

        titulo.setFont(new Font("Arial", Font.BOLD, 20));

        panelPrincipal.add(titulo, BorderLayout.NORTH);

        // TABLA.
        modelo = new DefaultTableModel(new String[]{
                "ID","NOMBRE","VALOR","FECHA"
                },
                0
        );

        tabla = new JTable(modelo);

        tabla.setRowHeight(25);

        JScrollPane scroll = new JScrollPane(tabla);

        panelPrincipal.add(scroll, BorderLayout.CENTER);

        // BOTONES.
        JPanel panelBotones = new JPanel(new FlowLayout(FlowLayout.CENTER, 10, 10));

        JButton btnConsultar = new JButton("CONSULTAR ORACLE");
        JButton btnGenerar = new JButton("GENERAR EXCEL");
        JButton btnAbrir = new JButton("ABRIR EXCEL");
        JButton btnLimpiar = new JButton("LIMPIAR");

        panelBotones.add(btnConsultar);
        panelBotones.add(btnGenerar);
        panelBotones.add(btnAbrir);
        panelBotones.add(btnLimpiar);

        panelPrincipal.add(panelBotones, BorderLayout.SOUTH);
        add(panelPrincipal);

        // EVENTOS.
        btnConsultar.addActionListener(e -> consultarOracle());
        btnGenerar.addActionListener(e -> generarExcel());
        btnAbrir.addActionListener(e -> abrirExcel());
        btnLimpiar.addActionListener(e -> limpiarTabla());
    }

    private void limpiarTabla() {
    }

    private void abrirExcel() {

    }

    private void generarExcel() {

    }

    private void consultarOracle() {

    }
}