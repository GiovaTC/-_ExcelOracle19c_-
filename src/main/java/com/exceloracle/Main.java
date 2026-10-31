package com.exceloracle;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
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

}