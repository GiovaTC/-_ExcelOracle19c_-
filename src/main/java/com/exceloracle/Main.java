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

    private Connection conectarOracle() throws SQLException {

        return DriverManager.getConnection(
                URL, USUARIO, CLAVE
        );
    }

    // CONSULTAR REGISTROS DE ORACLE!
    private void consultarOracle() {

        modelo.setRowCount(0);
        String sql = """
                SELECT ID, NOMBRE, VALOR, FECHA_REGISTRO FROM REGISTROS_EXCEL ORDER BY ID
                """;

        try (
                Connection conexion = conectarOracle();
                PreparedStatement ps = conexion.prepareStatement(sql);
                ResultSet rs = ps.executeQuery()
        ) {

            SimpleDateFormat formato = new SimpleDateFormat("yyyy-MM-dd");

            while (rs.next())
            {
                Object[] fila = {
                        rs.getInt("ID"),
                        rs.getString("NOMBRE"),
                        rs.getBigDecimal("VALOR"),

                        formato.format(rs.getDate("FECHA_REGISTRO"))
                };

                modelo.addRow(fila);
            }

            JOptionPane.showMessageDialog(
                    this,
                    "Consulta realizada correctamente.\n"
                            + "Registros: "
                            + modelo.getRowCount()
            );
        } catch (SQLException ex) {

            mostrarError(
                    "Error al consultar Oracle:\n"
                            + ex.getMessage()
            );
        }
    }

    // GENERAR ARCHIVO EXCEL!
    private void generarExcel()
    {
        if (modelo.getRowCount() == 0) {

            JOptionPane.showMessageDialog(
                    this,
                    "Primero consulta los registros de Oracle."
            );

            return;
        }

        JFileChooser selector =
                new JFileChooser();

        selector.setDialogTitle(
                "Guardar archivo Excel"
        );

        selector.setSelectedFile(
                new File("ReporteOracle19c.xlsx")
        );

        int resultado =
                selector.showSaveDialog(this);

        if (resultado !=
                JFileChooser.APPROVE_OPTION) {

            return;
        }

        File archivo =
                selector.getSelectedFile();

        // Asegurar extensión XLSX
        if (!archivo.getName()
                .toLowerCase()
                .endsWith(".xlsx")) {

            archivo = new File(
                    archivo.getAbsolutePath()
                            + ".xlsx"
            );
        }

        try (Workbook libro =
                     new XSSFWorkbook()) {

            Sheet hoja =
                    libro.createSheet(
                            "Reporte Oracle"
                    );

            // ESTILO DEL ENCABEZADO
            CellStyle estiloEncabezado =
                    libro.createCellStyle();

            org.apache.poi.ss.usermodel.Font fuente =
                    libro.createFont();

            fuente.setBold(true);

            fuente.setColor(
                    IndexedColors.WHITE.getIndex()
            );

            estiloEncabezado.setFont(fuente);

            estiloEncabezado
                    .setFillForegroundColor(
                            IndexedColors.DARK_BLUE
                                    .getIndex()
                    );

            estiloEncabezado
                    .setFillPattern(
                            FillPatternType
                                    .SOLID_FOREGROUND
                    );

            // CREAR ENCABEZADOS
            Row encabezado =
                    hoja.createRow(0);

            for (
                    int columna = 0;
                    columna < modelo.getColumnCount();
                    columna++
            ) {

                Cell celda =
                        encabezado.createCell(
                                columna
                        );

                celda.setCellValue(
                        modelo.getColumnName(
                                columna
                        )
                );

                celda.setCellStyle(
                        estiloEncabezado
                );
            }

            // CREAR FILAS
            for (
                    int fila = 0;
                    fila < modelo.getRowCount();
                    fila++
            ) {

                Row filaExcel =
                        hoja.createRow(
                                fila + 1
                        );

                for (
                        int columna = 0;
                        columna < modelo.getColumnCount();
                        columna++
                ) {

                    Cell celda =
                            filaExcel.createCell(
                                    columna
                            );

                    Object valor =
                            modelo.getValueAt(
                                    fila,
                                    columna
                            );

                    celda.setCellValue(
                            valor == null
                                    ? ""
                                    : valor.toString()
                    );
                }
            }

            // AJUSTAR ANCHO DE COLUMNAS
            for (
                    int columna = 0;
                    columna < modelo.getColumnCount();
                    columna++
            ) {

                hoja.autoSizeColumn(
                        columna
                );
            }

            // GUARDAR ARCHIVO
            try (
                    FileOutputStream salida =
                            new FileOutputStream(
                                    archivo
                            )
            ) {

                libro.write(salida);
            }

            JOptionPane.showMessageDialog(
                    this,
                    "Archivo Excel generado correctamente:\n"
                            + archivo.getAbsolutePath()
            );

            // VISUALIZAR EL ARCHIVO
            cargarExcelEnTabla(
                    archivo
            );

        } catch (Exception ex) {

            mostrarError(
                    "Error al generar Excel:\n"
                            + ex.getMessage()
            );
        }
    }

    private void abrirExcel() {

        JFileChooser selector = new JFileChooser();

        selector.setDialogTitle("SELECCIONAR archivo excel!");

        int resultado = selector.showOpenDialog(this);

        if ( resultado != JFileChooser.APPROVE_OPTION) {
            return;
        }

        File archivo = selector.getSelectedFile();
    }

    private void cargarExcelEnTabla(File archivo) {
        try (
                FileInputStream entrada = new FileInputStream(archivo);

                Workbook libro = WorkbookFactory.create(entrada)
            ) {

            Sheet hoja = libro.getSheetAt(0);
            modelo.setRowCount(0);
            modelo.setColumnCount(0);

            Row encabezado = hoja.getRow(0);

            if (encabezado == null)
            {
                throw new Exception("EL archivo excel " + "no tiene encabezado!");
            }

            // LEER ENCABEZADOS
            for (
                    int columna = 0;
                    columna < encabezado
                            .getLastCellNum();
                    columna++
            ) {

                Cell celda =
                        encabezado.getCell(
                                columna
                        );

                modelo.addColumn(
                        obtenerValorCelda(
                                celda
                        )
                );
            }

            // LEER FILAS
            for (
                    int fila = 1;
                    fila <= hoja.getLastRowNum();
                    fila++
            ) {

                Row filaExcel =
                        hoja.getRow(fila);

                if (filaExcel == null) {

                    continue;
                }

                Object[] datos =
                        new Object[
                                modelo.getColumnCount()
                                ];

                for (
                        int columna = 0;
                        columna < modelo
                                .getColumnCount();
                        columna++
                ) {

                    datos[columna] =
                            obtenerValorCelda(
                                    filaExcel.getCell(
                                            columna
                                    )
                            );
                }

                modelo.addRow(datos);
            }

            JOptionPane.showMessageDialog(
                    this,
                    "Archivo Excel cargado "
                            + "en la interfaz."
            );

        } catch (Exception ex) {

            mostrarError(
                    "Error al abrir Excel:\n"
                            + ex.getMessage()
            );
        }
    }

    private Object obtenerValorCelda(Cell celda)
    {
        if (celda == null)
        {
            return "";
        }

        DataFormatter formato = new DataFormatter();

        return formato.formatCellValue(celda);
    }


    private void limpiarTabla()
    {

        modelo.setRowCount(0);
    }

    private void mostrarError(String mensaje)
    {
        JOptionPane.showMessageDialog(this, mensaje, "Error", JOptionPane.ERROR_MESSAGE
        );
    }   

}