
#include "mainwindow.h"
#include "ui_mainwindow.h"

#include <QDir>
#include <QFile>
#include <QAction>
#include <QFileInfo>
#include <QFileDialog>
#include <QPushButton>
#include <QListWidget>
#include <QLineEdit>
#include <QProcessEnvironment>
#include <QHBoxLayout>
#include <QVBoxLayout>
#include <QDebug>
#include <QLibrary>
#include <QMessageBox>
#include <QProgressDialog>
#include <QSettings>
#include <QAxObject>
#include <QTextCodec>
#include <QDir>
#include <QDateTime>
#include <QMutex>

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);
//    this->setMinimumSize(400, 300);
    this->setWindowTitle(QStringLiteral("EXCEL&SQL"));
    this->setWindowIcon(QIcon(":/logo.ico"));

    row_start = -1;
    column_start = -1;
    row_count = -1;
    column_count = -1;
    _work_sheet = NULL;

    ui->label_output->clear();
    ui->label_info->setMinimumHeight(36);
    ui->label_type->setMinimumHeight(36);
    ui->label_version->setMinimumHeight(36);
    ui->label->setMinimumHeight(36);
    ui->comboBox->setMinimumHeight(36);
    ui->lineEdit_info->setMinimumHeight(36);
    ui->comboBox_type->setMinimumHeight(36);
    ui->lineEdit_version->setMinimumHeight(36);
    ui->pushButton_preview->setMinimumHeight(36);
    ui->pushButton_convertion->setMinimumHeight(36);

    ui->progressBar->setValue(0);

    initConfig = new QSettings("Convertion.ini", QSettings::IniFormat);
    initConfig->setIniCodec("GBK");


    menu = new QMenu();
    action = new QAction("删除", this);
    ui->listWidget->setContextMenuPolicy(Qt::CustomContextMenu);
    connect(action, SIGNAL(triggered()), this, SLOT(on_deleteRow()));


    ui->horizontalLayout_2->setAlignment(Qt::AlignRight);

    QStringList groupList = initConfig->childGroups();
    if(!groupList.isEmpty())
    {
//        QString group = groupList.at(0);
        foreach (QString title, groupList) {

            if(title.compare("MATCH") == 0){
                getInitValue(title, matchs);
                ui->comboBox->addItem("K歌比赛");
            }
            else if(title.compare("MEDIA") == 0){
                getInitValue(title, media);
                ui->comboBox->addItem("MV歌曲");
            }
            else if(title.compare("MP3") == 0){
                getInitValue(title, mp3);
                ui->comboBox->addItem("MP3歌曲");
            }
            else if(title.compare("ACTOR") == 0){
                getInitValue(title, actor);
                ui->comboBox->addItem("歌星");
            }
            else if(title.compare("SONGLIST") == 0){
                getInitValue(title, songlist);
                ui->comboBox->addItem("FM歌单");
            }
            else if(title.compare("MEDIALIST") == 0){
                getInitValue(title, medialist);
                ui->comboBox->addItem("排行旁");
            }
            else if(title.compare("HOT") == 0){
                getInitValue(title, hot);
                ui->comboBox->addItem("热度");
            }
            else if(title.compare("TYPE") == 0){
                getInitValue(title, type);

                ui->comboBox_type->addItems(type.values());
            }
        }
    }
}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::getInitValue(const QString &title, QMap<QString, QString> &value)
{
    initConfig->beginGroup(title);
    QStringList keyList=initConfig->childKeys();

    foreach(QString key,keyList)
    {
        QString str = initConfig->value(key).toString();

        value.insert(key, str);
    }
    initConfig->endGroup();
}

void MainWindow::errorMessageBox(const QString &error)
{

    ///表名或其他出错
    QString str = QString("%1操作错误\n 表名不正确，查看表名是否正确！！\n"
                           "未使用表是否删除。").arg(error);
    QMessageBox::warning(this, "异常出错", str);
}


void MainWindow::on_comboBox_currentIndexChanged(const QString &arg1)
{

}


void MainWindow::on_pushButton_preview_clicked()
{
    QString fileFormat(QStringLiteral("EXCEL文件(*.xlsx  *.xls  *.csv)"));
    QString document;
    defaultPath =  initConfig->value("DEFAULTPATH/path").toString();
    if(defaultPath.isEmpty())
        document = QProcessEnvironment::systemEnvironment().value("USERPROFILE")+"\\Desktop";
    else
        document = defaultPath;
    QStringList pathStrs = QFileDialog::getOpenFileNames(this,
                                                         QStringLiteral("视频转换"),
                                                         document,
                                                         fileFormat
                                                         );

    if(pathStrs.isEmpty())
        return;

    ui->listWidget->addItems(pathStrs);
}

void MainWindow::on_pushButton_convertion_clicked()
{
    ///初始化统计变量
    m_add = 0; m_update = 0; m_dele = 0;
    a_add = 0; a_update = 0; a_dele = 0;
    p3_add = 0; p3_update = 0; p3_dele = 0;
    match_update = 0; other_count = 0;
    if(isEmpty()){
        QMessageBox::warning(this, "提示", "版本不能是空！");
        return;
    }

    int count = ui->listWidget->count();
    for(int i=0; i<count; i++){
        QListWidgetItem *item = ui->listWidget->item(i);
        QString excelPath = item->text(); ;// = ui->lineEdit->text();
        if(excelPath.isEmpty())
            return;

        QDir dir;//(excelPath);
        dir.dirName();
        QString _path = excelPath;
        _path.replace(dir.dirName(), "");
        if(defaultPath.compare(_path) != 0)
        {
            initConfig->setValue("DEFAULTPATH/path", _path);
        }

        QString outPath = excelPath;
        if(outPath.indexOf(".xlsx") != -1) {
            outPath.replace(".xlsx", ".sql");
        }else{
            outPath.replace(".csv", ".sql");
        }

        if(i == 0){
            destFilePath = outPath;

            QFile file1(destFilePath);
            if(file1.exists()){
                file1.remove();
            }

            QFile file(destFilePath);
            if(file.open(QIODevice::Append | QIODevice::WriteOnly))
            {
                QTextStream out(&file);
                out << QString("USE `yiqiding_ktv`;\n");
            }
        }

        if(excelPath.indexOf(".csv") != -1)
        {
            QStringList list = excelPath.split("/");
            QString fileName = list.last();
            fileName = fileName.left(fileName.length() - 4);
            selectType(excelPath, outPath, fileName, true);

        } else {

            QAxObject *excel = NULL;
            QAxObject *work_books = NULL;
            QAxObject *work_book = NULL;
            excel = new QAxObject("Excel.Application");
            if (excel->isNull()) {//使用excel==NULL判断，是错误的
                QMessageBox::critical(0, "错误信息", "没有找到EXCEL应用程序");
                return;
            }

            excel->setProperty("Visible", false);
            work_books = excel->querySubObject("WorkBooks");


            work_books->dynamicCall("Open (const QString&)", QString(excelPath));
            QVariant title_value = excel->property("Caption");  //获取标题
            work_book = excel->querySubObject("ActiveWorkBook");
            QAxObject *work_sheets = work_book->querySubObject("WorkSheets");  //Sheets也可换用WorkSheets

            int sheet_count = work_sheets->property("Count").toInt();  //获取工作表数目
            int sheet_index = 0;

            for(int i=1; i<=sheet_count; i++)
            {
                _work_sheet = work_book->querySubObject("Sheets(int)", i);  //Sheets(int)也可换用Worksheets(int)
                QString work_sheet_name = _work_sheet->property("Name").toString();  //获取工作表名称

                if(work_sheet_name.isEmpty()){
                    continue;
                }

                sheet_index = i;
                if(sheet_index != 0)
                {
                    QAxObject *used_range = _work_sheet->querySubObject("UsedRange");
                    QAxObject *rows = used_range->querySubObject("Rows");
                    QAxObject *columns = used_range->querySubObject("Columns");
                    row_start = used_range->property("Row").toInt();  //获取起始行
                    column_start = used_range->property("Column").toInt();  //获取起始列
                    row_count = rows->property("Count").toInt();  //获取行数
                    column_count = columns->property("Count").toInt();  //获取列数

                    if(row_count < 0 || column_count < 0){
                        continue;
                    }

                    selectType(excelPath, outPath, work_sheet_name);
                }

            }

            excel->dynamicCall("Quit (void)");
        }

    }

    if(count > 0){
        setInfoText();
        insertLogSql();
        QMessageBox::information(NULL, "提示", "转换结束。");
    }
}

void MainWindow::on_deleteRow()
{
    QList<QListWidgetItem *> lists = ui->listWidget->selectedItems();
    for(int i=0; i<lists.count(); i++)
    {
        QListWidgetItem *item = ui->listWidget->takeItem(ui->listWidget->row(lists.at(i)));
        ui->listWidget->removeItemWidget(item);
    }
}

void MainWindow::selectType(const QString &srcPath, const QString &destPath,const QString &work_sheet_name,  const bool &isCsv)
{
    if(destFilePath.isEmpty())
        return;

    QFile file(destFilePath);
    if(!file.open(QIODevice::Append | QIODevice::WriteOnly))
    {
        QMessageBox::information(this, "提示", "文件打开失败！");
        return;
    }
    QTextStream out(&file);
    out << QString("/*%1*/\n").arg(work_sheet_name);

    ///加载数据到内存
    if (isCsv){
        if(!initRowValue_csv(srcPath)){
            return;
        }
    } else {
        if(!initRowValue_excel(_work_sheet)){
            return;
        }
    }

    this->setWindowTitle("脚本转换……");
    ui->progressBar->setRange(0, rowList.count()); //进度条

    ////type excel表名
    QString comboboxValue = ui->comboBox->currentText();   //
    if(matchs.keys().indexOf(work_sheet_name) != -1){
        match(out);
        match_update += rowList.count() - 1;
    }
    else if(media.keys().indexOf(work_sheet_name) != -1){
        if(work_sheet_name.compare("media") == 0){
            QString mediaInsertSql = QString("INSERT IGNORE INTO `media` "
                                             "(%1) VALUES ").arg(linkSql(fieldList));
            insertSqlStr(out, mediaInsertSql);
            m_add += rowList.count() - 1;
        }else if(work_sheet_name.compare("media_update_field") == 0
                 || work_sheet_name.compare("media_update_field1") == 0
                 || work_sheet_name.compare("media_update_field2") == 0
                 || work_sheet_name.compare("media_update_field3") == 0
                 || work_sheet_name.compare("media_update_field4") == 0
                 || work_sheet_name.compare("media_update_field5") == 0
                 ){

            QString updateSql = QString("UPDATE `media` SET ");
            updateSqlStr(out, updateSql);
            m_update += rowList.count() - 1;
        }
        else if(work_sheet_name.compare("media_delete") == 0){

            QString deleteSql = QString("DELETE FROM `media` WHERE `mid` ");
            deleteSqlStr(out, deleteSql);
            m_dele += rowList.count() - 1;
        }
        else
            errorMessageBox("MV歌曲");
    }
    else if(mp3.keys().indexOf(work_sheet_name) != -1){
        if(work_sheet_name.compare("mp3") == 0){
            QString mp3InsertSql = QString("INSERT IGNORE INTO `media_music` "
                                           "(%1) VALUES ").arg(linkSql(fieldList));
            insertSqlStr(out, mp3InsertSql);
            p3_add += rowList.count() - 1;
        }else if(work_sheet_name.compare("mp3_update_field") == 0
                 || work_sheet_name.compare("mp3_update_field1") == 0
                 || work_sheet_name.compare("mp3_update_field2") == 0
                 || work_sheet_name.compare("mp3_update_field3") == 0
                 || work_sheet_name.compare("mp3_update_field4") == 0
                 || work_sheet_name.compare("mp3_update_field5") == 0){

            QString updateSql = QString(" UPDATE `media_music` SET ");
            updateSqlStr(out, updateSql);
            p3_update += rowList.count() - 1;
        }else if(work_sheet_name.compare("mp3_delete") == 0){

            QString deleteSql = QString("DELETE FROM `media_music` WHERE `mmid` ");
            deleteSqlStr(out, deleteSql);
            p3_dele += rowList.count() - 1;
        }
        else
            errorMessageBox("MP3歌曲");
    }
    else if(actor.keys().indexOf(work_sheet_name) != -1){
        if(work_sheet_name.compare("actor") == 0){

//            QString actorInsertSql = QString("INSERT IGNORE INTO `actor` "
//                                           " (`sid`, `serial_id`, `name`,  `nation`, `sex`, "
//                                           " `pinyin`, `header`, `head`, `words`, `song_count`, "
//                                           " `stars`,`count`, `order`, `enabled`, `black`, "
//                                           " `info`) VALUES ");
            QString actorInsertSql = QString("INSERT IGNORE INTO `actor` "
                                           " (%1) VALUES ").arg(linkSql(fieldList));
            insertSqlStr(out, actorInsertSql);
            a_add += rowList.count() - 1;
        } else if(work_sheet_name.compare("actor_update_field") == 0
                  || work_sheet_name.compare("actor_update_field1") == 0
                  || work_sheet_name.compare("actor_update_field2") == 0
                  || work_sheet_name.compare("actor_update_field3") == 0
                  || work_sheet_name.compare("actor_update_field4") == 0
                  || work_sheet_name.compare("actor_update_field5") == 0){

            QString updateSql = QString("UPDATE `actor` SET ");
            updateSqlStr(out, updateSql);
            a_update += rowList.count() - 1;
        } else if(work_sheet_name.compare("actor_delete") == 0){

            QString deleteSql = QString("DELETE FROM `actor` WHERE `sid` ");
            deleteSqlStr(out, deleteSql);
            a_dele += rowList.count() - 1;
        }
        else
            errorMessageBox("歌星");
    }
    else if(songlist.keys().indexOf(work_sheet_name) != -1){
        if(work_sheet_name.compare("songlist") == 0){

            out << QString("DELETE FROM `songlist`;\n");
//            QString insertSql = QString("INSERT IGNORE INTO `songlist` "
//                                        " (`lid`, `serial_id`, `title`, `image`, `type`, `count`, `special`) VALUES "
//                                        );
            QString insertSql = QString("INSERT IGNORE INTO `songlist` "
                                        " (%1) VALUES ").arg(linkSql(fieldList));
            insertFMSqlStr(out, insertSql);
            other_count += rowList.count() - 1;
        }
        else if(work_sheet_name.compare("songlist_detail") == 0){

            out << QString("DELETE FROM `songlist_detail`;\n");
            QString insertSql = QString(" INSERT IGNORE INTO `songlist_detail`"
                                        "(%1) VALUES ").arg(linkSql(fieldList));
            insertSqlStr(out, insertSql);
            other_count += rowList.count() - 1;
        }
        else
            errorMessageBox("FM歌单");
    }
    else if(medialist.keys().indexOf(work_sheet_name) != -1){

        QString deleteSql = QString("DELETE FROM `media_list` where `type` = '%1';\n").arg(work_sheet_name);
        out << deleteSql;
        QString insertSql = QString("INSERT IGNORE INTO `media_list`(%1) VALUES ").arg(linkSql(fieldList));
        insertSqlStr(out, insertSql);
        other_count += rowList.count() - 1;

    } else if(hot.keys().indexOf(work_sheet_name) != -1){
        if(work_sheet_name.compare("hot_songs") == 0){

            QString updateSql = QString("UPDATE `media` SET count = count + %1 where mid = %2 ;");
            hotSqlStr(out, updateSql);
            other_count += rowList.count() - 1;
        }
        else if(work_sheet_name.compare("hot_singer") == 0){

            QString updateSql = QString("UPDATE `actor` SET count = count + %1 where sid = %2 ;");
            hotSqlStr(out, updateSql);
            other_count += rowList.count() - 1;
        }
        else
            errorMessageBox("热度");
    }////

    out << "\n";
    file.close();
}

QString MainWindow::linkSql(QStringList firstRow)
{
    QString retStr;
    for (int i=0; i<firstRow.size(); i++) {
        QString str = firstRow.at(i);
        retStr.append(QString("`%1`").arg(str));
        if (i != firstRow.size() - 1)
            retStr.append(", ");
    }

    return retStr;
}

void MainWindow::match(QTextStream &out)
{
    int valid_column = rowList.first().count();
    for(int i=1; i<rowList.count(); i++)
    {
        QStringList list = rowList.at(i);
        QString serial_idS, matchS;
        for(int j=0; j<valid_column; j++)
        {
            QString value = list.at(j);
            if(j == 0) {
                serial_idS = value;
            } else if(j == 1) {
                if(value.isEmpty() || value.compare("1") != 0)
                    matchS = "0";
                else
                    matchS = "1";
            }
        }

        if(serial_idS.isEmpty() || matchS.isEmpty())
            continue;

        QString insertSql = QString(" UPDATE `media` SET `match`='%1' WHERE `serial_id`='%2';\n")
                                    .arg(matchS)
                                    .arg(serial_idS);
        out << insertSql;

        ui->progressBar->setValue(i+1);
    }
}

QStringList MainWindow::splitCSVLine(const QString &lineStr)
{
    QStringList strList;
    QString str;

    int length = lineStr.length();
    int quoteCount = 0;
    int repeatQuoteCount = 0;

    for(int i = 0; i < length; ++i)
    {
        if(lineStr[i] != '\"')
        {
            repeatQuoteCount = 0;
            if(lineStr[i] != ',')
            {
                str.append(lineStr[i]);
            }
            else
            {
                if(quoteCount % 2)
                {
                    str.append(',');
                }
                else
                {
                    strList.append(str);
                    quoteCount = 0;
                    str.clear();
                }
            }
        }
        else
        {
            ++quoteCount;
            ++repeatQuoteCount;
            if(repeatQuoteCount == 4)
            {
                str.append('\"');
                repeatQuoteCount = 0;
                quoteCount -= 4;
            }
        }
    }
    strList.append(str);

    return qMove(strList);
}

void MainWindow::insertSqlStr(QTextStream &out, QString &sqlStr)
{
    int valid_column = rowList.at(0).count();
    for(int i=1; i<rowList.count(); i++)
    {    

        QString appSql;
        QStringList list = rowList.at(i);
        for(int j=0; j<valid_column; j++)
        {
            QString value = list.at(j);
            QString str;

            if(value.compare("NULL") == 0 || value.isEmpty()){
                str = QString(" NULL, ");
            } else {
                str = QString("'%1', ").arg(value);
            }
            appSql.append(str);
        }

        appSql.insert(0, "(");
        appSql.replace(appSql.length() - 2, 2, ' ');
        appSql.append(");");

        QString temp = sqlStr;
        temp.append(appSql);
        temp.append("\n");
        out << temp;

        ui->progressBar->setValue(i+1);
    }
}

void MainWindow::insertFMSqlStr(QTextStream &out, QString &sqlStr)
{
    int valid_column = rowList.at(0).count();
    for(int i=1; i<rowList.count(); i++)
    {
        QStringList list = rowList.at(i);
        QString lidS, serial_idS, titleS;
        lidS = list.at(0);
        serial_idS = list.at(1);
        titleS = list.at(2);

        QString appSql = QString("('%1', '%2', '%3', 'yqc', %4, 0);")
                .arg(lidS)
                .arg(serial_idS)
                .arg(titleS)
                .arg(QString("(select count(*) from songlist_detail where lid = %1)").arg(lidS));


        QString temp = sqlStr;
        temp.append(appSql);
        temp.append("\n");
        out << temp;

        ui->progressBar->setValue(i+1);
    }
}

void MainWindow::updateSqlStr(QTextStream &out, QString &sqlStr)
{
    ////要修改的字段
    QStringList fields = rowList.at(0);
    for(int i=1; i<rowList.count(); i++)
    {
        ///获取每列的值
        QStringList list = rowList.at(i);
        QStringList sql;
        for(int j=0; j<fields.count(); j++)
        {
            QString value = list.at(j);
            QString temp = value;
            if(temp.compare("NULL") != 0)
                temp = QString("'%1'").arg(value);

            sql.append(QString("`%1`=%2").arg(fields.at(j)).arg(temp));
        }

        ///拼接SQL SET语句
        QString tempUpdate = sqlStr;
        for(int i = 1; i<sql.count(); i++)
        {
            tempUpdate.append(sql.at(i));

            if(i != sql.size() - 1){
                tempUpdate.append(", ");
            }
        }

        ///添加WHERE语句
        tempUpdate.append(" WHERE ");
        tempUpdate.append(sql.at(0));
        tempUpdate.append(";");

        tempUpdate.append("\n");
        out << tempUpdate;

        ui->progressBar->setValue(i+1);
    }
}

void MainWindow::deleteSqlStr(QTextStream &out, QString &sqlStr)
{
    int valid_column = rowList.at(0).count();
    for(int i=1; i<rowList.count(); i++)
    {
        QStringList list = rowList.at(i);
        QString midS;
        for(int j=0; j<valid_column; j++){

            if(j == 0){
                midS = list.at(j);
            }
        }

        if(midS.isEmpty()){
            continue;
        }

        QString deleteSql = QString("%1 = %2 ;\n").arg(sqlStr).arg(midS);
        out << deleteSql;

        ui->progressBar->setValue(i+1);
    }
}

void MainWindow::hotSqlStr(QTextStream &out, QString &sqlStr)
{
    ////要修改的字段
    QStringList fields = rowList.at(0);
    for(int i=1; i<rowList.count(); i++)
    {
        ///获取每列的值
        QStringList list = rowList.at(i);
        QString midStr;
        QString countStr;
        for(int j=0; j<fields.count(); j++){

            QString value  = list.at(j);
            if(j == 0) midStr = value;
            if(j == 1) countStr = value;
        }

        QString updateSql = QString(sqlStr).arg(countStr).arg(midStr);
        updateSql.append("\n");
        out << updateSql;

        ui->progressBar->setValue(i+1);
    }
}

bool MainWindow::initRowValue_excel(QAxObject *work_sheet)
{
    rowList.clear();
    this->setWindowTitle("加载EXCEL数据……");
    ui->progressBar->setRange(0, row_count); //进度条
    QString sheetName = _work_sheet->property("Name").toString();
    for(int i=row_start; i<=row_count; i++)
    {
        QStringList row;
        for(int j=column_start; j<=column_count; j++)
        {
            QString value =  getCellValue(work_sheet, i, j);
            if(value.isEmpty()){
                value = "NULL";
            }

            row.append(value);
        }

        ///是否是空行
        if(!row.isEmpty()){
            rowList.append(row);
        }

        ///表头字段
        if(i == row_start){
            fieldList.clear();
            fieldList = row;
        } else {

            ///错误筛选
            QString text;
            foreach (QString str, row) {
                text.append(str).append("  ");
            }

            ui->label_output->setText(QString("行：%1 ").arg(i));
            int column = -1;
            if (sheetName.compare("media") == 0
                    || sheetName.compare("media_update_field") == 0
                    || sheetName.compare("media_update_field1") == 0
                    || sheetName.compare("media_update_field2") == 0
                    || sheetName.compare("media_update_field3") == 0
                    || sheetName.compare("media_update_field4") == 0
                    || sheetName.compare("media_update_field5") == 0
                    ){
                column = analyzeRowData(row);

            } else if (sheetName.compare("actor") == 0
                       || sheetName.compare("actor_update_field") == 0
                       || sheetName.compare("actor_update_field1") == 0
                       || sheetName.compare("actor_update_field2") == 0
                       || sheetName.compare("actor_update_field3") == 0
                       || sheetName.compare("actor_update_field4") == 0
                       || sheetName.compare("actor_update_field5") == 0
                       ){
                column = analyzeActorRowData(row);
            }

            if(column != -1){
                QMessageBox::warning(this, "错误提示",
                                     QString("错误表：%4 \n"
                                             "行：%1   列：%2  值：%3")
                                     .arg(i).arg(column+1).arg(QString(row.at(column))).arg(sheetName));
                return false;
            }

        }
        ui->progressBar->setValue(i+1);
    }

    return true;
}

bool MainWindow::initRowValue_csv(const QString &path)
{
    QFile file(path);
    QStringList CSVList;
    CSVList.clear();
    rowList.clear();
    if(file.open(QIODevice::ReadOnly))
    {
        bool isfirst = true;
        QTextStream steam(&file);

//        qDebug() << " start time : " << QDateTime::currentDateTime().toString("yyyy-MM-dd hh:mm:ss:zzz");
        this->setWindowTitle("加载CSV数据……");
        while (!steam.atEnd()) {

            CSVList = splitCSVLine(steam.readLine());
            if(CSVList.at(0).isEmpty()){
                continue;
            }
            rowList.append(CSVList);


            if(isfirst){

                isfirst = false;
                fieldList.clear();
                fieldList = CSVList;
            }
        }
//        qDebug() << " end time : " << QDateTime::currentDateTime().toString("yyyy-MM-dd hh:mm:ss:zzz");
    }

    int row = -1, column = -1;
    bool status = true;
    QString fileName = path.split("/").last();
    fileName.remove(".csv");
    if (fileName.compare("media") == 0
        || fileName.compare("media_update_field") == 0
        || fileName.compare("media_update_field1") == 0
        || fileName.compare("media_update_field2") == 0
        || fileName.compare("media_update_field3") == 0
        || fileName.compare("media_update_field4") == 0
        || fileName.compare("media_update_field5") == 0
       ){


        ui->progressBar->setRange(0, rowList.count());
        for(int i=0; i<rowList.count(); i++){

            QString text;
            foreach (QString str, rowList.at(i)) {
                text.append(str).append("  ");
            }
            ui->label_output->setText(QString("行：%1").arg(i+1));
            column = analyzeRowData(rowList.at(i));
            if(column != -1){
                row = i;
                status = false;
                break;
            }

            ui->progressBar->setValue(i+1);
        }
    } else if (fileName.compare("actor") == 0
               || fileName.compare("actor_update_field") == 0
               || fileName.compare("actor_update_field1") == 0
               || fileName.compare("actor_update_field2") == 0
               || fileName.compare("actor_update_field3") == 0
               || fileName.compare("actor_update_field4") == 0
               || fileName.compare("actor_update_field5") == 0
               ){

        ui->progressBar->setRange(0, rowList.count());
        for(int i=0; i<rowList.count(); i++)
        {
            QString text;
            foreach (QString str, rowList.at(i)) {
                text.append(str).append("  ");
            }
            ui->label_output->setText(QString("行：%1").arg(i+1));
            column = analyzeActorRowData(rowList.at(i));
            if(column != -1){
                row = i;
                status = false;
                break;
            }
            ui->progressBar->setValue(i+1);
        }
    }

    if(!status){
        QMessageBox::warning(this, "错误提示",
                             QString("错误行：%1   列：%2")
                             .arg(row+1).arg(column+1));
    }

    return status;
}

QString MainWindow::getCellValue(QAxObject *work_sheet, int row, int column)
{
    QAxObject *cell = work_sheet->querySubObject("Cells(int,int)", row, column);
    QVariant cell_value = cell->property("Value");  //获取单元格内容
    QString value = cell_value.toString();
    value = value.simplified();

    return value;
}

bool MainWindow::isStringLength(const QString &str, const int &length)
{
    ///字符个数
//    QTextCodec *gbk = QTextCodec::codecForName("GBK");
//    QByteArray _src = gbk->fromUnicode(str);

    if(str.length() > length){
        return true;
    }

    return false;
}

bool MainWindow::isNumber(const QString &str)
{
    for (int i=0; i<str.size(); i++)
    {
        if(str.at(i) < '0' || str.at(i) > '9')
            return false;
    }
    return true;
}

int MainWindow::analyzeRowData(const QStringList &row)
{
    int ret = -1;
    for(int i=0; i<fieldList.count(); i++){

        QString value = row.at(i);
        if(value.isEmpty()){
            ret = i;
            break;
        }

        if ( fieldList.at(i).compare("mid") == 0
             || fieldList.at(i).compare("count") == 0){

            if(!queryNumberAndLength(value, 10)){
                 ret = i;
                 break;
             }
        } else if ( fieldList.at(i).compare("serial_id") == 0
                    || fieldList.at(i).compare("artist_sid_1") == 0
                    || fieldList.at(i).compare("artist_sid_2") == 0){

            if(!queryNumberAndLength(value, 10, true)){
                ret = i;
                break;
            }

        } else if ( fieldList.at(i).compare("name") == 0
                    || fieldList.at(i).compare("pinyin") == 0){

            if(isStringLength(value, 200)){
                ret = i;
                break;
            }

        } else if ( fieldList.at(i).compare("language") == 0
                    || fieldList.at(i).compare("type") == 0){

             if(!queryNumberAndLength(value, 2)){
                 ret = i;
                 break;
             }

        } else if ( fieldList.at(i).compare("singer") == 0
                    || fieldList.at(i).compare("header") == 0
                    || fieldList.at(i).compare("path") == 0){

            if(isStringLength(value, 100)){
                ret = i;
                break;
            }

        } else if (fieldList.at(i).compare("head") == 0){

            if(isStringLength(value, 1)){
                ret = i;
                break;
            }
        } else if (fieldList.at(i).compare("words") == 0){

             if(!queryNumberAndLength(value, 3)){
                 ret = i;
                 break;
             }
        } else if ( fieldList.at(i).compare("original_track") == 0
                    || fieldList.at(i).compare("sound_track") == 0){

            if(value.compare("0") != 0 && value.compare("1") != 0){
                ret = i;
                break;
            }
        } else if (fieldList.at(i).compare("start_volume_1") == 0
                   || fieldList.at(i).compare("start_volume_2") == 0){

            if(!queryNumberAndLength(value, 3, true)){
                ret = i;
                break;
            }
        } else if (fieldList.at(i).compare("prelude") == 0){

            if(!queryNumberAndLength(value, 3, true)){
                ret = i;
                break;
            }

        } else if (fieldList.at(i).compare("effect") == 0
                   || fieldList.at(i).compare("version") == 0){

            if(!queryNumberAndLength(value, 2, true)){
                ret = i;
                break;
            }
        } else if (fieldList.at(i).compare("stars") == 0){

        } else if (fieldList.at(i).compare("enabled") == 0
                   || fieldList.at(i).compare("black") == 0
                   || fieldList.at(i).compare("hot") == 0){

            if(value.compare("0") != 0 && value.compare("1") != 0){
                ret = i;
                break;
            }
        } else if (fieldList.at(i).compare("match") == 0){

            if(value.compare("NULL") != 0){
            if(value.compare("0") != 0 && value.compare("1") != 0){
                ret = i;
                break;
            }
            }
        } else if ( fieldList.at(i).compare("resolution") == 0
                    || fieldList.at(i).compare("quality") == 0
                    || fieldList.at(i).compare("source") == 0
                    || fieldList.at(i).compare("rhythm") == 0
                    || fieldList.at(i).compare("pitch") == 0
                    || fieldList.at(i).compare("lang_part") == 0){

            if(!queryNumberAndLength(value, 2, true)){
                ret = i;
                break;
            }
        } else if (fieldList.at(i).compare("desc_count") == 0){

        }

    }

    return ret;
}

int MainWindow::analyzeActorRowData(const QStringList &row)
{
    int ret = -1;
    for(int i=0; i<fieldList.count(); i++){

        QString value = row.at(i);
        if(value.isEmpty()){
            ret = i;
            return ret;
        }
        if ( fieldList.at(i).compare("sid") == 0
             || fieldList.at(i).compare("serial_id") == 0
             || fieldList.at(i).compare("count") == 0){

            if(!queryNumberAndLength(value, 10)){
                 ret = i;
                 break;
             }
        } else if (fieldList.at(i).compare("name") == 0){

            if(isStringLength(value, 200)){
                ret = i;
                break;
            }

        } else if (fieldList.at(i).compare("nation") == 0){

             if(!queryNumberAndLength(value, 2)){
                 ret = i;
                 break;
             }

        } else if (fieldList.at(i).compare("sex") == 0){

            if(!queryNumberAndLength(value, 1)){
                ret = i;
                break;
            }

        } else if (fieldList.at(i).compare("pinyin") == 0){

            if(isStringLength(value, 200)){
                ret = i;
                break;
            }

        } else if (fieldList.at(i).compare("header") == 0){

            if(isStringLength(value, 100)){
                ret = i;
                break;
            }
        } else if ( fieldList.at(i).compare("head") == 0 ){

            if(isStringLength(value, 1)){
                ret = i;
                break;
            }
        } else if (fieldList.at(i).compare("words") == 0){

             if(!queryNumberAndLength(value, 3)){
                 ret = i;
                 break;
             }
        } else if (fieldList.at(i).compare("song_count") == 0){

            if(!queryNumberAndLength(value, 5)){
                ret = i;
                break;
            }
        } else if (fieldList.at(i).compare("stars") == 0){

        } else if (fieldList.at(i).compare("order") == 0){

            if(!queryNumberAndLength(value, 5, true)){
                ret = i;
                break;
            }
        } else if (fieldList.at(i).compare("enabled") == 0 || fieldList.at(i).compare("black") == 0){

            if(value.compare("0") != 0 && value.compare("1") != 0){
                ret = i;
                break;
            }
        }
    }

    return ret;
}

bool MainWindow::queryNumberAndLength(const QString &value, const int &length, const bool &isNull)
{
    if(isNull){

        if(value.compare("NULL") != 0){

            if(!isNumber(value) || isStringLength(value, length)){
                 return false;
             }
        }
    } else {

        if(!isNumber(value) || isStringLength(value, length)){
             return false;
         }
    }

    return true;
}

void MainWindow::insertLogSql()
{
    QFile file(destFilePath);
    if(!file.open(QIODevice::Append | QIODevice::WriteOnly))
    {
        QMessageBox::information(this, "提示", "文件打开失败！");
        return;
    }

    QTextStream out(&file);
    QString sqlStr = QString("insert into `update_sql_log`"
                             "(`type`, `version`, `create_time`, `update_time`, `info`) "
                             "VALUES('%1', '%2', '%3', NOW(), '%4');")
                            .arg(type.key(ui->comboBox_type->currentText()))
                            .arg(ui->lineEdit_version->text())
                            .arg(QDateTime::currentDateTime().toString("yyyy-MM-dd hh:mm:ss"))
                            .arg(ui->lineEdit_info->text());

    out << QString("/*update sql log*/\n");
    out << sqlStr << "\n";

    file.close();
}

void MainWindow::setInfoText()
{
    QString temp;
    if(m_add != 0)
        temp.append(QString("新增歌曲%1; ").arg(m_add));
    if(m_update != 0)
        temp.append(QString("更新歌曲%1; ").arg(m_update));
    if(m_dele != 0)
        temp.append(QString("删除歌曲%1; ").arg(m_dele));


    if(a_add != 0)
        temp.append(QString("新增歌星%1; ").arg(a_add));
    if(a_update != 0)
        temp.append(QString("更新歌星%1; ").arg(a_update));
    if(a_dele != 0)
        temp.append(QString("删除歌星%1; ").arg(a_dele));

    if(p3_add != 0)
        temp.append(QString("新增MP3歌曲%1; ").arg(p3_add));
    if(p3_update != 0)
        temp.append(QString("更新MP3歌曲%1; ").arg(p3_update));
    if(p3_dele != 0)
        temp.append(QString("删除MP3歌曲%1; ").arg(p3_dele));

    if(match_update != 0)
        temp.append(QString("更新K歌歌曲%1; ").arg(match_update));

    if(other_count != 0)
        temp.append(QString("其他%1; ").arg(other_count));

    ui->lineEdit_info->setText(temp);
}


bool MainWindow::isEmpty()
{
    if(ui->lineEdit_version->text().isEmpty()){
        return true;
    }

    return false;
}

void MainWindow::on_listWidget_customContextMenuRequested(const QPoint &pos)
{
    menu->clear();

    menu->addAction(action);
    menu->exec(QCursor::pos());
}


void MainWindow::setStyleSheet()
{
    QString style(" QPushButton{\
          background-color: rgb(255, 255, 255);\
          border: 1px solid rgb(170, 170, 170);\
          color:rgb(18, 18, 18);\
          font-size:14px;\
          border-radius:5px;\
      }\
      QPushButton:hover{\
          background-color:rgb(255, 255, 255);\
          border: 1px solid rgb(42, 42, 42);\
          color:rgb(42, 42, 42);\
      }\
      \
      QPushButton:pressed{\
          background-color: rgb(255, 146, 62);\
          border: 1px solid rgb(255, 146, 62);\
          color:rgb(255, 255, 255);\
      } \
     \
     QLineEdit{\
         border: 1px solid rgb(170, 170, 170);\
         color:rgb(202, 202, 202);\
         border-radius:5px;\
     }\
     QLineEdit:hover{\
         border: 1px solid rgb(42, 42, 42);\
         color:rgb(202, 202, 202);\
     }\
     QLineEdit:pressed{\
         border: 1px solid rgb(255, 146, 62);\
         color:rgb(88, 88, 88);\
     }\
     QLineEdit:disabled{\
         color:rgb(202, 202, 202);\
         border: 1px solid rgb(170, 170, 170);\
     }\
     QListWidget{\
         background-color:rgb(247, 246, 246);\
         alternate-background-color:rgb(234, 234, 234);\
         font-size:14px;\
     }\
     QListWidget::item{\
         height:40;\
     }\
     ");

//     this->setStyleSheet(style);

     }
