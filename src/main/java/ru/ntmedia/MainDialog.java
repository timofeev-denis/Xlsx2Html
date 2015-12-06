package ru.ntmedia;

import javax.swing.*;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;
import java.awt.event.*;
import java.io.File;
import java.net.URISyntaxException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Locale;

public class MainDialog extends JDialog {
    private JPanel contentPane;
    private JButton buttonOK;
    private JButton buttonCancel;
    private JTextField srcTextField;
    private JButton setSrcFolderButton;
    private JTextField destTextField;
    private JButton setDestFolderButton;
    private JSeparator separator;
    //private String srcFolder;
    //private String destFolder;

    public MainDialog() {
        setContentPane(contentPane);
        setModal(true);
        getRootPane().setDefaultButton(buttonOK);
        setComponentsText();
        buttonOK.setEnabled(false);

        buttonOK.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                onOK();
            }
        });

        buttonCancel.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                onCancel();
            }
        });

// call onCancel() when cross is clicked
        setDefaultCloseOperation(DO_NOTHING_ON_CLOSE);
        addWindowListener(new WindowAdapter() {
            public void windowClosing(WindowEvent e) {
                onCancel();
            }
        });

// call onCancel() on ESCAPE
        contentPane.registerKeyboardAction(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                onCancel();
            }
        }, KeyStroke.getKeyStroke(KeyEvent.VK_ESCAPE, 0), JComponent.WHEN_ANCESTOR_OF_FOCUSED_COMPONENT);

        setSrcFolderButton.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent actionEvent) {
                JFileChooser fileDialog = new JFileChooser();
                fileDialog.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
                if( fileDialog.showOpenDialog(MainDialog.this) == JFileChooser.APPROVE_OPTION ) {
                    srcTextField.setText(fileDialog.getSelectedFile().toString());
                }
            }
        });
        setDestFolderButton.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent actionEvent) {
                JFileChooser fileDialog = new JFileChooser();
                fileDialog.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
                if( fileDialog.showSaveDialog(MainDialog.this) == JFileChooser.APPROVE_OPTION ) {
                    destTextField.setText(fileDialog.getSelectedFile().toString());
                }
            }
        });
        srcTextField.getDocument().addDocumentListener(new DocumentListener() {
            public void insertUpdate(DocumentEvent documentEvent) {
                updateOkButton();
            }

            public void removeUpdate(DocumentEvent documentEvent) {
                updateOkButton();
            }

            public void changedUpdate(DocumentEvent documentEvent) {
                updateOkButton();
            }
        });
        destTextField.getDocument().addDocumentListener(new DocumentListener() {
            public void insertUpdate(DocumentEvent documentEvent) {
                updateOkButton();
            }

            public void removeUpdate(DocumentEvent documentEvent) {
                updateOkButton();
            }

            public void changedUpdate(DocumentEvent documentEvent) {
                updateOkButton();
            }
        });
    }

    private void updateOkButton() {
        if(srcTextField.getText().equals("") || destTextField.getText().equals("")) {
            buttonOK.setEnabled(false);
        } else {
            buttonOK.setEnabled(true);
        }
    }

    private void onOK() {
        //System.err.println(srcTextField.getText());
        if(!Files.exists(Paths.get(srcTextField.getText()))) {
            JOptionPane.showMessageDialog(null, "Указанный каталог с файлами Excel не существует. Выберите другой каталог.", "Конвертация файлов", JOptionPane.INFORMATION_MESSAGE);
            return;
        }
        if(!Files.exists(Paths.get(destTextField.getText()))) {
            JOptionPane.showMessageDialog(null, "Указанный каталог для сохранения файлов HTML не существует. Выберите другой каталог.", "Конвертация файлов", JOptionPane.INFORMATION_MESSAGE);
            return;
        }
        App app = new App(srcTextField.getText(), destTextField.getText());
        try {
            app.getTemplateCfg(getClass().getProtectionDomain().getCodeSource().getLocation().toURI().getPath());
        } catch (URISyntaxException e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(null, "Не удалось определить каталог с шаблонами.", "Конвертация файлов", JOptionPane.INFORMATION_MESSAGE);
            return;
        }
        app.convertAllFiles();
        dispose();
    }

    private void onCancel() {
// add your code here if necessary
        dispose();
    }

    private void setComponentsText() {
        UIManager.put("FileChooser.openDialogTitleText", "Открыть");
        UIManager.put("FileChooser.saveDialogTitleText", "Сохранить");
        UIManager.put("FileChooser.lookInLabelText", "Каталог");
        UIManager.put("FileChooser.openButtonText", "Открыть");
        UIManager.put("FileChooser.saveButtonText", "Выбрать");
        UIManager.put("FileChooser.cancelButtonText", "Отмена");
        UIManager.put("FileChooser.fileNameLabelText", "Имя файла");
        UIManager.put("FileChooser.folderNameLabelText", "Каталог");
        UIManager.put("FileChooser.filesOfTypeLabelText", "Типы файлов");
        UIManager.put("FileChooser.openButtonToolTipText", "OpenSelectedFile");
        UIManager.put("FileChooser.cancelButtonToolTipText","Отмена");
        UIManager.put("FileChooser.fileNameHeaderText","Имя файла");
        UIManager.put("FileChooser.upFolderToolTipText", "UpOneLevel");
        UIManager.put("FileChooser.homeFolderToolTipText","Desktop");
        UIManager.put("FileChooser.newFolderToolTipText","CreateNewFolder");
        UIManager.put("FileChooser.listViewButtonToolTipText","List");
        UIManager.put("FileChooser.newFolderButtonText","CreateNewFolder");
        UIManager.put("FileChooser.renameFileButtonText", "RenameFile");
        UIManager.put("FileChooser.deleteFileButtonText", "DeleteFile");
        UIManager.put("FileChooser.filterLabelText", "Типы файлов");
        UIManager.put("FileChooser.detailsViewButtonToolTipText", "Details");
        UIManager.put("FileChooser.fileSizeHeaderText","Size");
        UIManager.put("FileChooser.fileDateHeaderText", "DateModified");
        UIManager.put("FileChooser.acceptAllFileFilterText", "Все файлы");
    }
}
