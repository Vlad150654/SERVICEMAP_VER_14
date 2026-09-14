unit cfCommon;

interface

uses
  Winapi.Messages, System.Classes, System.SysUtils,

  ComCtrls,
  System.RegularExpressions,
  Winapi.Windows,
  System.IniFiles,
  Vcl.Forms,

  OXmlUtils, OXmlPDOM, OXmlPSeq,
  MapBasic_INT,
  csCommon,
  Common,
  csGDALOGR,
  csGDALDataSet;

Type
  PCostFactory = ^TCostFactory;
  TCostFactory = Record//ценовые факторы
    id : Integer;//порядковый номер
    Table : String;//название таблицы
    Field : String;//поле в таблице MapInfo
    Alias : String;//поле на русском
    //101 зоны транспортного безразличия по объекту ztb
    //102 внутризональное_расстояние по объекту
//    CalcType : Integer;//тип расчета; 0 - нет расчета 1 по дорогам 2 по прямой НЕ АКТУАЛЬНО!!!
    fType : String;//тип поля
    Link : Integer;//индекс в массиве cfZU для занесения рачета подъезда или граница объекта (ztb или igs)НЕ АКТУАЛЬНО!!!
    gko  : Integer;//признак расчета гос.кад.оценки
  end;

  TPartitionData = Record//ценовые факторы
    idn : String;//IDN from Sourse
    id : Byte;//порядковый номер AnyData
    table : String;//название таблицы
    value : String;//значение поля
    area : Double;
    overlap  : Double;//площадь перекрытия
  end;

  TNodeList = Record
    node : String;
    pos : Integer;//позиция узла или первого аттрибута при узле
    check : Array of Integer;//номера колонок для проверки значений
    attr : StringArray;
  end;

var
  neednode : Array of TNodeList;
  needattr : StringArray;
  posattr : Array of Integer;//позиции вставки аттрибутов

Const
  NUMBER_MESSAGE = WM_USER + 103;//Сообщение должно обладать идентификатором в диапазоне от WM_USER+100 до Ç7FFFF

 UserAgent : Array [0..150] of String =
    ('Mozilla / 5.0 (Windows; U; Windows NT 5.0; fr-FR; rv: 1.7.1) Gecko / 20040707',
    'Mozilla / 5.0 (X11; U; Linux i686; fr-FR; rv: 1.7.3) Gecko / 20040913',
    'Mozilla / 5.0 (Windows; U; WinNT4.0; de-AT; rv: 1.2) Gecko / 20021126',
    'Mozilla / 5.0 (X11; U; Linux i686; en-US; rv: 1.3b) Gecko / 20030125',
    'Mozilla / 5.0 (X11; U; Linux i586; en-US; rv: 0.9.9) Gecko / 20020513',
    'Mozilla / 5.0 (X11; U; Linux i686; rv: 1.7.8) Gecko / 20050927 Debian / 1.7.8-1sarge3',
    'Mozilla / 5.0 (Windows; U; Windows NT 5.0; fr; rv: 1.7.8) Gecko / 20050511',
    'Mozilla / 5.0 (X11; U; Linux i686; en-US; rv: 1.8.0.1) Gecko / 20060126',
    'Mozilla / 5.0 (X11; U; Linux i586; en-US; rv: 1.2) Gecko / 20050223',
    'Mozilla / 5.0 (Windows; U; Windows NT 5.0; en-US; rv: 1.2a) Gecko / 20020910',
    'Mozilla / 5.0 (Windows; U; WinNT4.0; en-US; rv: 1.2b) Gecko / 20021016',
    'Mozilla / 5.0 (Windows; U; Windows NT 5.1; en-US; rv: 1.7.3) Gecko / 20040910',
    'Mozilla / 5.0 (X11; U; Linux i686; de-AT; rv: 1.7.2) Gecko / 20040810 Debian / 1.7.2-2',
    'Mozilla / 5.0 (Windows; U; Windows NT 5.1; en-US; rv: 1.5) Gecko / 20030916',
    'Mozilla / 5.0 (Windows; U; Windows NT 5.0; en-US; rv: 1.0rc2) Gecko / 20020510',
    'Mozilla / 5.0 (X11; U; Linux i686; en-US; rv: 1.7.8) Gecko / 20050831 Debian / 1.7.8-1sarge2',
    'Mozilla / 5.0 (Windows; U; Windows NT 5.0; de-AT; rv: 1.3a) Gecko / 20021212',
    'Mozilla / 5.0 (Macintosh; U; PPC Mac OS X; en-US; rv: 1.2b) Gecko / 20021016',
    'Mozilla / 5.0 (Windows; U; WinNT4.0; en-US; rv: 1.0.0) Gecko / 20020530',
    'Mozilla / 5.0 (X11; U; Linux i686; en-US; rv: 1.0.1) Gecko / 20020826',
    'Mozilla / 5.0 (Windows; U; Windows NT 5.0; en-US; rv: 1.3) Gecko / 20030312',
    'Mozilla / 5.0 (Windows; U; Windows NT 5.1; de-AT; rv: 1.7.11) Gecko / 20050728',
    'Mozilla / 5.0 (X11; U; SunOS sun4u; en-US; rv: 1.3) Gecko / 20030318',
    'Mozilla / 5.0 (Windows; U; Windows NT 5.0; en-US; rv: 1.1) Gecko / 20020826',
    'Mozilla / 5.0 (Windows; U; Windows NT 5.1; de-AT; rv: 1.8a5) Gecko / 20041122',
    'Mozilla / 5.0 (X11; U; Linux i686; de-AT; rv: 1.7.2) Gecko / 20040804',
    'Mozilla / 5.0 (Windows; U; Windows NT 5.1; en-US; rv: 1.5) Gecko / 20031007',
    'Mozilla / 5.0 (X11; U; Linux i686; de-AT; rv: 1.3) Gecko / 20030430 Debian / 1.3-5',
    'Mozilla / 5.0 (Windows; suburl; Windows NT 5.1; de-AT; rv: 1.8a1) Gecko / 20040520',
    'Mozilla / 5.0 (Windows; U; Windows NT 5.0; it-IT; rv: 1.7.12) Gecko / 20050915',
    'Mozilla / 5.0 (Windows; U; Windows NT 5.1; en-US; rv: 1.8a2) Gecko / 20040704',
    'Mozilla / 5.0 (Windows; U; Windows NT 5.1; de; rv: 1.8) Gecko / 20051111',
    'Mozilla / 5.0 (Windows; U; Windows NT 5.0; en-US; rv: 1.7.7) Gecko / 20050414',
    'Mozilla / 5.0 (Windows;; Windows NT 5.1; rv: 1.7.2) Gecko / 20040804',
    'Mozilla / 5.0 (X11; U; Linux i686; fr-FR; rv: 1.5) Gecko / 20031007',
    'Mozilla / 5.0 (Windows; U; Windows NT 5.0; en-US; rv: 1.0rc3) Gecko / 20020523',
    'Mozilla / 5.0 (Windows; U; Windows NT 5.1; en-US; rv: 1.7.8) Gecko / 20050511',
    'Mozilla / 5.0 (X11; U; Linux i686; en-US; rv: 1.6) Gecko / 20040413 Debian / 1.6-5',
    'Mozilla / 5.0 (X11; U; Linux i686; en-US; rv: 1.7.13) Gecko / 20060417',
    'Mozilla / 5.0 (X11; U; OpenBSD i386; en-US; rv: 1.7.13) Gecko / 20060901',
    'Mozilla / 5.0 (Windows; U; Windows NT 5.1; de-AT; rv: 1.8a1) Gecko / 20040520',
    'Mozilla / 5.0 (X11; U; Linux i686; en-US; rv: 1.7.8) Gecko / 20050921',
    'Mozilla / 5.0 (X11; U; Linux i686; en-US; rv: 1.7.8) Gecko / 20050511',
    'Mozilla / 5.0 (Windows; U; Windows NT 5.1; ru-RU; rv: 1.3) Gecko / 20030309',
    'Mozilla / 5.0 (Windows; U; Windows NT 5.1; en-US; rv: 1.7) Gecko / 20040616',
    'Mozilla / 5.0 (X11; U; Linux i686; uk-UA; rv: 0.9.6+) Gecko / 20011122',
    'Mozilla / 5.0 (X11; U; Linux x86_64; fr; rv: 1.6) Gecko / 20040115',
    'Mozilla / 5.0 (X11; U; Linux i686; en-US; rv: 0.9.3) Gecko / 20010801',
    'Mozilla / 5.0 (X11; U; Linux i686; en-US; rv: 1.4) Gecko / 20030624',
    'Mozilla / 5.0 (Windows NT 10.0; WOW64; rv: 39.0) Gecko / 20100101 Firefox / 80.0',
    'Mozilla / 5.0 (Windows NT 10.0; Win64; x64; rv: 80.0) Gecko / 20100101 OspGg8Fz-27 Firefox / 80.0',
    'Mozilla / 5.0 (Windows NT 10.0; Win64; x64; rv: 81.0) Gecko / 20100101 Firefox / 80.0',
    'Mozilla / 5.0 (Macintosh; Intel Mac OS X 10.14; rv: 81.0) Gecko / 20100101 Firefox / 80.0',
    'Mozilla / 5.0 (Windows NT 6.1; Win64; x64; rv: 80.0) Gecko / 20100101 Firefox / 80.0 / Qb43BjQg-56',
    'Mozilla / 5.0 (Windows NT 6.1; Win64; x64; rv: 80.0) Gecko / 20100101 VJ6nP2zX-53 Firefox / 80.0',
    'Mozilla / 5.0 (Windows NT 6.1; Win64; x64; rv: 80.0) Gecko / 20100101 B5zjdD2e-51 Firefox / 80.0',
    'Mozilla / 5.0 (Windows NT 6.3; Win64; x64; rv: 80.0) Gecko / 20100101 htH7cfqb-28 Firefox / 80.0',
    'Mozilla / 5.0 (X11; U; Linux x86_64; en-US; rv: 80.0) Gecko / 20100101 Kubuntu / 20.04 (Intrepid) Firefox / 80.0',
    'Mozilla / 5.0 (Windows NT 6.1; Win64; x64; rv: 80.0) Gecko / 20100101 Myjd9FM4-29 Firefox / 80.0',
    'Mozilla / 5.0 (Windows NT 6.2; Win64; x64; rv: 80.0) Gecko / 20100101 Firefox / 80.0, gzip (gfe)',
    'Mozilla / 5.0 (Windows NT 10.0; rv: 80.0) Gecko / 20100101 null Firefox / 80.0',
    'Mozilla / 5.0 (X11; FreeBSD powerpc; rv: 80.0) Gecko / 20100101 Firefox / 80.0',
    'Mozilla / 5.0 (Windows NT 10.0; Win64; x64; rv: 80.0) Gecko / 20100101 Firefox / 80.0 / 8mqNJauL-25',
    'Mozilla / 5.0 (Windows NT 6.1; rv: 80.0) Gecko / 20100101 Firefox / 80.0 / 10csO9CgK-99',
    'Mozilla / 5.0 (Windows NT 6.2; rv: 80.0.1) Gecko / 20100101 Firefox / 80.0.1 анонимизировано Abelssoft 1254735417',
    'Mozilla / 5.0 (Windows NT 10.0; WOW64; x64; rv: 80.0) Gecko / 20100101 Firefox / 80.0',
    'Mozilla / 5.0 (Windows NT 6.1; WOW64; rv: 54.0) Gecko / 20100101 Firefox / 80.0 / 7rpJ8CqK-00',
    'Mozilla / 5.0 (Windows NT 10.0; Win64; x64; rv: 80.0.1) Gecko / 20200505 Firefox / 80.0.1',
    'Mozilla / 5.0 (Windows NT 6.2; rv: 80.0.1) Gecko / 20100101 Firefox / 80.0.1 анонимно с помощью Abelssoft 1024581946',
    'Mozilla / 5.0 (Windows NT 10.0; Win64; x64; rv: 80.0) Gecko / 20100101 VhsBnoJO-36 Firefox / 80.0',
    'Mozilla / 5.0 (Windows NT 10.0; Win64; x64; rv: 80.0) Gecko / 20100101 yHL3o1a7-15 Firefox / 80.0',
    'Mozilla / 5.0 (Windows NT 10.0; Win64; x64; rv: 80.0) Gecko / 20100101 lecsIgCc-34 Firefox / 80.0',
    'Mozilla / 5.0 (Windows NT 10.0; rv: 80.0) Gecko / 20100101 Firefox / 80.0 анонимизировано Abelssoft 1663170808',
    'Mozilla / 5.0 (Windows NT 10.0; Win64; x64; rv: 80.0) Gecko / 20100101 nWQexAmN-33 Firefox / 80.0',
    'Mozilla / 5.0 (Windows NT 10.0; rv: 80.0) Gecko / 20100101 Firefox / 80.0 анонимизировано Abelssoft 865842111',
    'Mozilla / 5.0 (Windows NT 6.3; Win64; x64; rv: 80.0; Tesseract / 1.0) Gecko / 20100101 Firefox / 80.0',
    'Mozilla / 5.0 (Windows NT 6.1; Win64; x64; rv: 80.0) Gecko / 20100101 l3FNS2i0-1 Firefox / 80.0',
    'Mozilla / 5.0 (Windows NT 6.1; Win64; x64; rv: 80.0) Gecko / 20100101 l3FNS2i0-1 l3FNS2i0-1 Firefox / 80.0',
    'Mozilla / 5.0 (Windows NT 10.0; Win64; x64; rv: 80.0) Gecko / 20100101 Firefox / 80.0 IAV / FF80',
    'Mozilla / 5.0 (Windows NT 10.0; Win64; x64; rv: 80.0) Gecko / 20100101 Firefox / 80.0 / 8mqErPuL-62',
    'Mozilla / 5.0 (Windows NT 10.0; Win64; x64; rv: 80.0) Gecko / 20100101 Firefox / 80.0 / 8mqSsSuL-81',
    'Mozilla / 5.0 (Windows NT 10.0; Win64; x64; rv: 80.0) Gecko / 20100101 lQHGDnYu-22 Firefox / 80.0',
    'Firefox / 80.0.1 (x64 de); анонимно Abelssoft 674940923',
    'Mozilla / 5.0 (Macintoh; Intel Mac OS X 10.15; rv: 80.01) Gecko / 2s0100101 Firefox / 80.01',
    'Mozilla / 5.0 (Windows NT 10.0; Win64; x64; rv: 80.0) Gecko / 20100101 Firefox / 80.0 / mE2P7b4C-5',
    'Mozilla / 5.0 (Windows NT 6.3; Win64; x64; rv: 80.0) Gecko / 20100101 Firefox / 80.0 / 7rpJ8CqK-00',
    'Mozilla / 5.0 (Windows NT 6.1; Win64; x64; rv: 80.0) Gecko / 20100101 Firefox / 80.0 / 8mqMlPuL-96',
    'Mozilla / 5.0 (Windows NT 10.0; Win64; x64; rv: 80.0) Gecko / 20100101 Firefox / 80.0 / 8mqFuKuL-88',
    'Mozilla / 5.0 (Windows NT 10.0; Win64; x64; rv: 80.0) Gecko / 20100101 Firefox / 80.0 / 4erS7ObW-66',
    'Mozilla / 5.0 (Windows NT 10.0; Win64; x64; rv: 80.0) Gecko / 20100101 YfCmF2SI-6 Firefox / 80.0',
    'Mozilla / 5.0 (X11; Linux ppc64le; rv: 80.0) Gecko / 20100101 Firefox / 80.0',
    'Mozilla / 5.0 (X11; Fedora; Linux x86_64; rv: 80.0.1) Gecko / 20200505 Firefox / 80.0.1',
    'Mozilla / 5.0 (Windows NT 10.0; Win64; x64; rv: 80.0) Gecko / 20100101 Firefox / 80.0 Контейнер приложений',
    'Mozilla / 5.0 (Windows NT 6.1; rv: 80.0) Gecko / 20100101 Firefox / 80.0 анонимизировано Abelssoft 469857876',
    'Mozilla / 5.0 (Windows NT 6.1; rv: 80.0; Tesseract / 1.0) Gecko / 20100101 Firefox / 80.0',
    'Mozilla / 5.0 (X11; U; Linux x86_64; rv: 80.0) Gecko / 20100101 Firefox / 80.0',
    'Mozilla / 5.0 (Windows NT 6.1; Win64; x64; rv: 80.0) Gecko / 20100101 Firefox / 80.0 / 7rpJ8CqK-00',
    'Mozilla / 5.0 (Windows NT 10.0; rv: 80.0) Gecko / 20100101 Firefox / 80.0 анонимизировано Abelssoft 1541784439',
    'Mozilla / 5.0 (Windows NT 10.0; rv: 80.0) Gecko / 20100101 Firefox / 80.0 анонимно с помощью Abelssoft 1024581946',
    'Mozilla / 5.0 (Windows NT 6.2; rv: 80.0.1) Gecko / 20100101 Firefox / 80.0.1 анонимизировано Abelssoft 350370323',
    'Mozilla / 5.0 (Windows NT 6.1; Win64; x64; rv: 80.0) Gecko / 20100101 Firefox / 80.0 / 8mqPtVuL-9',
    'Mozilla / 5.0 (Windows NT 6.1; Win64; x64; rv: 80.0) Gecko / 20100101 Firefox / 80.0 / 8mqNiKuL-34',
    'Mozilla / 5.0 (Windows NT 6.1; Win64; x64; rv: 80.0) Gecko / 20100101 Firefox / 80.0 / 8mqFuKuL-88',
    'Mozilla / 5.0 (Windows NT 6.1; Win64; x64; rv: 80.0) Gecko / 20100101 Firefox / 80.0 / 8mqAoGuL-19',
    'Mozilla / 5.0 (Windows NT 6.1; Win64; x64; rv: 80.0) Gecko / 20100101 Firefox / 80.0 / 8mqSsSuL-81',
    'Mozilla / 5.0 (Windows NT 6.1; Win64; x64; rv: 80.0) Gecko / 20100101 Firefox / 80.0 / 8mqDiPuL-36',
    'Mozilla / 5.0 (Windows NT 6.1; Win64; x64; rv: 80.0) Gecko / 20100101 Firefox / 80.0 / 8mqNiKuL-99',
    'Mozilla / 5.0 (Windows NT 6.1; Win64; x64; rv: 80.0) Gecko / 20100101 Firefox / 80.0 / 8mqQaZuL-16',
    'Mozilla / 5.0 (Windows NT 10.0; Win64; x64; rv: 80.0) Gecko / 20100101 Firefox / 80.0 / 8mqULwuL-67',
    'Mozilla / 5.0 (X11; Linux x86_64: 80.0) Gecko / 20100101 Firefox / 80.0',
    'Mozilla / 5.0 (Windows NT 10.0; Win64; x64; rv: 80.0) Gecko / 20100101 Firefox / 80.0 / pyi1hrhO-1',
    'Mozilla / 5.0 (Windows NT 6.3; Win64; x64; rv: 80.0) Gecko / 20100101 Firefox / 80.0 / 9uiP7EnX-09',
    'Mozilla / 5.0 (Windows NT 6.2; Win64; x64; rv: 80.0) Gecko / 20100101 Firefox / 80.0 hinqyu8',
    'Mozilla / 5.0 (Windows NT 6.2; Win64; x64; rv: 80.0) Gecko / 20100101 Firefox / 80.0 m7k806',
    'Mozilla / 5.0 (Windows NT 6.2; Win64; x64; rv: 80.0) Gecko / 20100101 Firefox / 80.0 rkcdu4',
    'Mozilla / 5.0 (Windows NT 6.2; Win64; x64; rv: 80.0) Gecko / 20100101 Firefox / 80.0 g2rx5zf2',
    'Mozilla / 5.0 (Windows NT 6.2; Win64; x64; rv: 80.0) Gecko / 20100101 Firefox / 80.0 ukymizvpxq6',
    'Mozilla / 5.0 (Windows NT 6.2; Win64; x64; rv: 80.0) Gecko / 20100101 Firefox / 80.0 r5e98',
    'Mozilla / 5.0 (Windows NT 6.2; Win64; x64; rv: 80.0) Gecko / 20100101 Firefox / 80.0 l182zvolvw6',
    'Mozilla / 5.0 (Windows NT 6.2; Win64; x64; rv: 80.0) Gecko / 20100101 Firefox / 80.0 tkx1yjj66',
    'Mozilla / 5.0 (Windows NT 6.2; Win64; x64; rv: 80.0) Gecko / 20100101 Firefox / 80.0 r0qkezrsr0',
    'Mozilla / 5.0 (Windows NT 10.0; Win64; rv: 86.1) Gecko / 20100101 Firefox / 86.1 Gecko / 20100101 Firefox / 80.0',
    'Mozilla / 5.0 (Windows NT 6.2; rv: 80.0.1) Gecko / 20100101 Firefox / 80.0.1 анонимизировано Abelssoft 1855802700',
    'Mozilla / 5.0 (Windows NT 10.0; rv: 80.0) Gecko / 20100101 Firefox / 80.0 анонимизировано Abelssoft 1964569802',
    'Mozilla / 5.0 (Windows NT 10.0; Win64; x64; rv: 80.0) Gecko / 20100101 Firefox / 80.0 / hfJvl00D-38',
    'Mozilla / 5.0 (Windows NT 10.0; Win64; x64; rv: 80.0) Gecko / 20100101 hfJvl00D-38 Firefox / 80.0 / hfJvl00D-38',
    'Mozilla / 5.0 (Windows NT 6.1; Win64; x64; rv: 80.0) Gecko / 20100101 null Firefox / 80.0',
    'Mozilla / 5.0 (Windows NT 10.0; Win64; x64; rv: 80.0) Gecko / 20100101 Firefox / 80.0 / 8mqPaQuL-91',
    'Mozilla / 5.0 (Windows NT 10.0; Win64; x64; rv: 80.0) Gecko / 20100101 Firefox / 80.0 / 8mqDoRuL-01',
    'Mozilla / 5.0 (Windows NT 10.0; Win64; x64; rv: 80.0) Gecko / 20100101 Firefox / 80.0 / 8mqNiKuL-34',
    'Mozilla / 5.0 (Windows NT 6.1; WOW64; rv: 80.0) Gecko / 20100101 Firefox / 80.0',
    'Mozilla / 5.0 (Windows NT 10.0; Win64; x64; rv: 80.0) Gecko / 20100101 Firefox / 80.0 Edge / 17.17134',
    'Mozilla / 5.0 (X11; Ubuntu; Linux i686; rv: 80.0) Gecko / 20100101 Firefox / 80.0, gzip (gfe)',
    'Mozilla / 5.0 (Windows NT 10.0; Win64; x64; rv: 80.0; Tesseract / 1.0) Gecko / 20100101 Firefox / 80.0',
    'Mozilla / 5.0 (Windows NT 10.0; WOW64; rv: 80.0) Gecko / 20100101 Firefox / 80.0',
    'Mozilla / 5.0 (Windows NT 6.1; Win64; x64; rv: 80.0; Tesseract / 1.0) Gecko / 20100101 Firefox / 80.0',
    'Mozilla / 5.0 (Windows NT 6.1; Win64; x64; rv: 80.0) Gecko / 20100101 SJF9yVvE-18 Firefox / 80.0',
    'Mozilla / 5.0 (Windows NT 10.0; Win64; x64; rv: 80.0) Gecko / 20100101 Firefox / 80.0 / 8mqAoGuL-19',
    'Mozilla / 5.0 (Windows NT 10.0; Win64; x64; rv: 80.0) Gecko / 20100101 Firefox / 80.0 / 8mqEzXuL-46',
    'Mozilla / 5.0 (Windows NT 10.0; Win64; x64; rv: 70.0) Gecko / 20100101 Firefox / 80.0',
    'Mozilla / 5.0 (Windows NT 6.3; Win64; x64; rv: 80.0) Gecko / 20100101 Aq52zRuQ-32 Aq52zRuQ-32 Firefox / 80.0',
    'Mozilla / 5.0 (Windows NT 10.0; Win64; x64; rv: 80.0) Gecko / 20100101 Firefox / 80.0 / 7rpJ8CqK-00',
    'Mozilla / 5.0 (X11; OpenBSD amd64; rv: 80.0) Gecko / 20100101 Firefox / 80.0',
    'Mozilla / 5.0 (Windows NT 10.0; Win64; x64; rv: 80.0) Gecko / 20100101 Firefox / 80.0 / 8mqQaZuL-16',
    'Mozilla / 5.0 (Windows NT 6.3; Win64; x64; rv: 80.0) Gecko / 20100101 Aq52zRuQ-32 Firefox / 80.0',
    'Mozilla / 5.0 (Windows NT 10.0; Win64; x64; rv: 80.0) Gecko / 20100101 Firefox / 80.0 / 2kcL6ZrR-71',
    'Mozilla / 5.0 (X11; Arch Linux i686; rv: 2.0) Gecko / 20110321 Firefox / 4.0',
    'Opera / 9.80 (X11; Linux i686; Ubuntu / 14.10) Presto / 2.12.388 Версия / 12.16.2',
    'Opera / 9.80 (X11; Linux i686; Ubuntu / 14.10) Presto / 2.12.388 Версия / 12.16',
    'Opera / 9.80 (Macintosh; Intel Mac OS X 10.14.1) Presto / 2.12.388 Версия / 12.16',
    'Opera / 9.80 (Windows NT 6.0) Presto / 2.12.388 Версия / 12.14');

//procedure ReadOXmlPSeq(filename : String; var sl : TStringList);
procedure ReadOXmlPDOM(filename, rootname : String; var log, sl : TStringList);
function ExamenOGR(var stBar: TStatusBar; filename, temppath : String) : Boolean;

implementation

//procedure ExtractNode(offset : String; node : OXmlPDOM.PXMLNode; var sl : TStringList); overload
//var
//  nod : PXMLNode;
//begin
//  if node.HasAttributes then
//  begin
//    nod := node.FirstAttribute;
//    while Assigned(nod) do
//    begin
////      sl.Add(format('%s@ %s : %s', [offset, nod.NodeName, nod.NodeValue]));
//      nod := nod.NextSibling;
//    end;
//  end;
//
//  if node.HasChildNodes then
//  begin
//    nod := node.FirstChild;
//    while Assigned(nod) do
//    begin
//      if nod.NodeType = OXmlUtils.ntElement then //ntText - текст при узле
//      begin
//        if (nod.FirstChild <> nil) and (nod.FirstChild.NodeType = OXmlUtils.ntText) then
//          sl.Add(format('%s# %s: %s', [offset, nod.NodeName, nod.FirstChild.Text]))
//        else
//          sl.Add(format('%s# %s', [offset, nod.NodeName]));
//        ExtractNode(offset + '  ', nod, sl);
//      end;
//      nod := nod.NextSibling;
//    end;
//  end;
//end;
//
(*
procedure ExtractNode(node : OXmlPDOM.PXMLNode; var sl : TStringList; var nodelist : StringArray); overload
var
  nod : PXMLNode;
  i : Integer;
begin
  if node.HasAttributes then
  begin
    nod := node.FirstAttribute;
    while Assigned(nod) do
    begin
      for i := Low(neednode) to High(neednode) do
        if CText(nod.NodeName, neednode[i]) then
        begin
          nodelist[i] := nod.NodeValue;
          break;
        end;
      nod := nod.NextSibling;
    end;
  end;

  if node.HasChildNodes then
  begin
    nod := node.FirstChild;
    while Assigned(nod) do
    begin
      if nod.NodeType = OXmlUtils.ntElement then //ntText - текст при узле
      begin
        if (nod.FirstChild <> nil) and (nod.FirstChild.NodeType = OXmlUtils.ntText) then
        for i := Low(neednode) to High(neednode) do
          if CText(nod.NodeName, neednode[i]) then
          begin
            nodelist[i] := nod.FirstChild.Text;
            break;
          end;

//        if (nod.FirstChild <> nil) and (nod.FirstChild.NodeType = OXmlUtils.ntText) then
//          sl.Add(format('%s# %s: %s', [delimiter, nod.NodeName, nod.FirstChild.Text]))
//        else
//          sl.Add(format('%s# %s', [delimiter, nod.NodeName]));
        ExtractNode(nod, sl, nodelist);
      end;
      nod := nod.NextSibling;
    end;
  end;
end;      *)
//https://github.com/beNative/lazarus/blob/master/components/OXml/demos/Lazarus/uXmlTest.pas#L940
(*procedure ReadOXmlPSeq(filename : String; var sl : TStringList);
var
  s : String;
  i : Integer;
var
  xSeq: OXmlPSeq.TXMLSeqParser;
  xDataIsOpen: Boolean;
  xRowNode, nod: OXmlPDOM.PXMLNode;
  nodelist : StringArray;

begin
//filename:='D:\Delphi\Comps\DevExpress18_2_3\ExpressQuantumGrid\Demos\Data\Companies.xml';
  sl.Clear;
  xSeq := OXmlPSeq.TXMLSeqParser.Create;
  try
    xSeq.InitFile(filename);
    if xSeq.GoToPath('/ListForRating/Objects/Parcels') then//находим группу
      xSeq.SkipNextChildElementHeader(xDataIsOpen{%H-});
    if xDataIsOpen then
    begin
      while xSeq.ReadNextChildNode({%H-}xRowNode) do
      begin
        SetLength(nodelist, Length(neednode));
//        sl.Add(format(%s - %s', [xRowNode.NodeName, xRowNode.NodeValue]));
        if xRowNode.FindChild('Parcel', nod) then //поиск конкретного узла
          _Navigate(nod, sl, nodelist);
        s := '';
        for i := 0 to Length(neednode) do
          s := s + nodelist[i] + '~';
        sl.Add(Copy(s, 1, Length(s)-1));
        SetLength(nodelist, 0);
      end;
      xSeq.SkipNextChildElementHeader(xDataIsOpen{%H-});
    end;

    xSeq.GoToPath('/');//go to end
  finally
    xSeq.Free;
  end;
end;
*)
////////////////////////////////////////////
// ntElement, ntAttribute, ntText
procedure _Navigate(const aNode: OXmlPDOM.PXMLNode; parentnodename : String;
  var nodelist : StringArray);
var
  nod : OXmlPDOM.PXMLNode;
  i, j : Integer;
  nodename, dat : String;
begin
  if aNode.HasAttributes then//first=Construction of others
  begin
    nod := aNode.FirstAttribute;
    while Assigned(nod) and (nod.ParentNode.NodeName = parentnodename) do
    begin
      for i := Low(needattr) to High(needattr) do
        if CText(nod.NodeName, needattr[i]) then
        begin
          dat := nod.NodeValue;

//          mc := reg.Matches(dat, '\d\d\d\d-\d\d-\d\d', [roIgnoreCase, roMultiLine]);
          if {(mc.count > 0) and} (Pos('Date', needattr[i]) > 0) then
            SetLength(dat, 10);

          if nodelist[posattr[i]] = '' then
            nodelist[posattr[i]] := dat//nod.NodeValue
          else
            nodelist[posattr[i]] := nodelist[posattr[i]] + ',' + dat;//nod.NodeValue;
          break;
        end;
      nod := nod.NextSibling;
    end;
  end;

  if aNode.HasChildNodes then
  begin
    nod := aNode.FirstChild;
    while Assigned(nod) do
    begin
      if nod.NodeType = OXmlUtils.ntElement then
      begin
        for i := Low(neednode) to High(neednode) do
          if CText(nod.NodeName, neednode[i].node) then
          begin//ntText - ищем текст при узле
           if (nod.FirstChild <> nil) and (nod.FirstChild.NodeType = OXmlUtils.ntText) then
             if nodelist[neednode[i].pos] = '' then//1-й узел по этому названию
               nodelist[neednode[i].pos] := nod.text;
            nodename := neednode[i].node;//для проверки аттрибутов по родителю
            needattr := neednode[i].attr;
            SetLength(posattr, Length(needattr));
            for j := Low(posattr) to High(posattr) do
              posattr[j] := neednode[i].pos + j;
            break;
          end;
        _Navigate(nod, nodename, nodelist);
      end;
      nod := nod.NextSibling;
    end;
  end;
end;
//Самый быстрый и удобный для памяти DOM для Pascal.
procedure ReadOXmlPDOM(filename, rootname : String; var log, sl : TStringList);
var
  xXML : OXmlPDOM.IXMLDocument;
  root, node : OXmlPDOM.PXMLNode;
  nodelist : StringArray;//список всех узлов и аттрибутов
  i, k, m : Integer;
  s, nodename : String;
  sl2: TStringList;
begin
  xXML := OXmlPDOM.CreateXMLDoc;
  try
    xXML.LoadFromFile(filename);
    root := xXML.DocumentElement; //returns the root node(first element in the document)
    nodename := Copy(rootname, 1, Length(rootname)-1);
    if root.FindChild('ListInfo', node) then //находим группу
      if root.FindChild('Objects', node) then
        if node.FindChild(rootname, node) then
          if node.FindChild(nodename, node) then
          while Assigned(node) do
          begin
            k :=  0;
            for i := Low(neednode) to High(neednode) do
            begin
              if Length(neednode[i].attr) > 0 then
                Inc(k, Length(neednode[i].attr))
              else
                Inc(k, 1);
            end;
            SetLength(nodelist, k);

            needattr := neednode[0].attr;
            SetLength(posattr, Length(needattr));
            for i := Low(posattr) to High(posattr) do
              posattr[i] := neednode[0].pos + i;
            _Navigate(node, neednode[0].node, nodelist);//первым делом проверяем атрибуты

            s := '';
            for i := Low(nodelist) to High(nodelist) do
            begin
              m := 0;
              for k := 1 to length(nodelist[i]) do
                if nodelist[i][k] = '"' then
                  inc(m);
              if m mod 2 <> 0 then
              begin
                log.Add(format('файл: %s%s  КН %s Удален символ " в строке: %s:',
                  [filename, #10, nodelist[0], nodelist[i]]));
                nodelist[i] := StringReplace(nodelist[i], '"', '', [rfReplaceAll]);
              end;

              if (Pos(#10, nodelist[i]) > 0) or (Pos(#13, nodelist[i]) > 0) then
              begin
                log.Add(format('файл: %s%s  КН %s Удалены символы возврата каретки и перевода строк',
                  [filename, #10, nodelist[0]]));
                nodelist[i] := StringReplace(nodelist[i], #13, '', [rfReplaceAll]);
                nodelist[i] := StringReplace(nodelist[i], #10, '', [rfReplaceAll]);
              end;

              if Pos(#9, nodelist[i]) > 0 then
              begin
                log.Add(format('файл: %s%s  КН %s Удален символ табуляции',
                  [filename, #10, nodelist[0]]));
                nodelist[i] := StringReplace(nodelist[i], #9, '', [rfReplaceAll]);
              end;

              if Length(nodelist[i]) > 32750 then
              begin
                log.Add(format('файл: %s%s  КН %s Информация обрезана: %s',
                  [filename, #10, nodelist[0], nodelist[i]]));
                SetLength(nodelist[i], 32750);
              end;
              s := s + Trim(nodelist[i]) + '~';
            end;

            sl.Add(Copy(s, 1, Length(s)-1));
            SetLength(nodelist, 0);
            node := node.NextSibling;
          end;
  finally
    SetLength(needattr, 0);
    SetLength(posattr, 0);
    xXML := nil;
  end;
end;

function ExamenOGR(var stBar: TStatusBar; filename, temppath : String) : Boolean;
var
  ds : TcsOGRDataSet;
  i, j, k : Integer;
  fieldname, tab, path, err, s, s2, cn, x, y : String;
  sl, sl2, sl0 : TStringList;
  ini: TMemIniFile;
  mc : TMatchCollection;

  procedure Info(msg : String);
  begin
    stBar.Panels[1].Text := msg;
    stBar.Repaint;
    Application.ProcessMessages;
  end;

  function num_polygons(i : Integer) : Integer;//число полигонов
  begin
    Result := 1;
    While i < sl2.Count do
    begin
      if sl2[i][1] = '(' then
        Result := Result + 1
      else
        exit;
      Inc(i);
    end;
  end;

  function findpart(s : String; var j : Integer) : Boolean;//поиск части в коллекции
  begin
    Result := true;
    j := 0;
    While j < sl2.Count do
    begin
      if Pos(s, sl2[j]) > 0 then
        exit;
      Inc(j);
    end;
    Result := False;
  end;

  function points(s : String) : String;//набор точек
  var
    i, k : Integer;
  begin
    mc := reg.Matches(s, '\(\d.+\d\)', [roMultiLine]);
    s := StringReplace(mc[0].Value, ',', ')(', [rfReplaceAll]);
    s := StringReplace(s, ' ', ',', [rfReplaceAll]);
    k := 0;
    for i := 1 to Length(s) do
      if s[i] = ',' then
        Inc(k);
    Result := IntToStr(k) + s + ' ';
  end;

begin
  Result := true;
  tab := miEval('PathToTableName$("%s")', [filename]);
  if not TableExists(tab) then
    miDo('Open table "%s"', [filename]);
  fieldname := FindCadNumByName(tab);
  if fieldname = '' then
    exit;

  miDo('Commit Table %s As "%sGDALtest.TAB" TYPE NATIVE Charset "WindowsCyrillic"'+
    ' CoordSys Earth Projection 1, 104', [tab, TempPath]);

  path := GetAppFolder + 'GDAL\';
  SetDllDirectory(PWideChar(path));
  csGDALInit(PWideChar(path));
  sl0 := TStringList.Create;
  sl := TStringList.Create;
  sl2 := TStringList.Create;
  ds := TcsOGRDataSet.Create(nil);
  ini := TMemIniFile.Create(GetAppFolder + 'ServiceMap.ini');

  try
    ini.ReadSectionValues('Translate', sl0);
    ds.TablePath := TempPath + 'GDALtest.TAB';
    ds.Open;
    ds.First;//это обязательно, иначе первая запись не будет прочитана
    k := 0;
    while not ds.Eof do
    begin
      Inc(k);
      cn := ds.FieldByName(fieldname).asString;
      Info(' Проверка графики ' + cn);
      if not ds.IsGeometryValid(err) then
      begin
        sl.Add(cn + ' ' + err);
//         sl.Add('NOTVAL '+ ds.GetWKTGeometry(False));
        sl.Add(cn + ' ' + 'MAKEVAL ' + ds.GetWKTGeometry(True));
      end;
//        else sl.Add('VAL '+ds.GetWKTGeometry(False));
      ds.Next;
    end;

    if sl.Count > 0 then
    begin
      path := ExtractFilePath(filename) + 'Error\';
      ExShellExecute('cmd', PChar(format('/C rd /S /Q "%s"', [path])), SW_HIDE, true, 0);
      ForceDirectories(path);
      SafeCloseTable('ErrorTable');
      miDo('Create table ErrorTable (КН Char(35), _errors Char(90)) File "%sErrorTable"',[path]);
      miDo('Create Map For ErrorTable CoordSys Earth Projection 1, 104');
      miDo('Set CoordSys table ErrorTable');
      for i := 0 to sl.Count-1 do
      begin
        s := sl[i];
        mc := reg.Matches(s, '\d\d:\d\d:\d{1,7}:\d{1,}', []);
        if mc.Count = 1 then
        begin
//          cn := mc.Item[0].Value;
          k := Pos(' ', s);
          cn := Copy(s, 1, k-1);
          Info(' Создание таблици ошибок ' + cn);
          s := Trim(Copy(s, Length(cn)+1, Length(s)));
          k := Pos('MAKEVAL', s);
          if k = 0 then
          begin
            mc := reg.Matches(s, '\D+', []);
            s2 := Trim(mc.Item[0].Value);
            s := Trim(Copy(s, Length(s2)+1, Length(s)));
            j := Pos(' ', s);
            x := Trim(Copy(s, 1, j-1));
            y := Trim(Copy(s, j+1, Length(s)));
            s2 := sl0.Values[s2];
            if s2 = '' then//не нашли оставляем без перевода
              s2 := mc.Item[0].Value;
            miDo('Insert into ErrorTable (obj,КН,_errors)'+
            ' Values(CreatePoint(%s,%s), "%s", "%s")', [x, y, cn, s2]);
          end else

          begin
            j := 0;
            s := StringReplace(Trim(Copy(s, k+8, Length(s))), '),', ')' + lf, [rfReplaceAll]);
            sl2.Text := s;

            if Pos('LINESTRING', sl2[0]) = 1 then
            begin
              s := 'Create pline into Variable ob ';
              s := s + points(sl2[j]);
              for j := j+1 to sl2.Count-1 do
                if sl2[j][1] = '(' then
                  s := s + points(sl2[j])
                else
                  break;
              miDo(s);
              miDo('Insert into ErrorTable (obj,КН,_errors) Values(ob, "%s", "%s")',
                [cn, 'Новый объекта']);
            end else

            if (Pos('POLYGON', sl2[0]) = 1) or (Pos('MULTIPOLYGON', sl2[0]) = 1) then
            begin
              s := format('Create Region into Variable ob %d ', [num_polygons(1)]);
              for j := 0 to sl2.Count-1 do
                s := s + points(sl2[j]);
              miDo(s);
              miDo('Insert into ErrorTable (obj,КН,_errors) Values(ob, "%s", "%s")',
                [cn, 'Новый объект']);
            end else

            if Pos('MULTILINESTRING', sl2[0]) = 1 then
            begin
              s := format('Create pline into Variable ob multiple %d ', [num_polygons(j+1)]);
              s := s + points(sl2[j]);
              for j := j+1 to sl2.Count-1 do
                if sl2[j][1] = '(' then
                  s := s + points(sl2[j])
                else
                  break;
              miDo(s);
              miDo('Insert into ErrorTable (obj,КН,_errors) Values(ob, "%s", "%s")',
                [cn, 'Новый объект']);
            end else

            if Pos('GEOMETRYCOLLECTION', sl2[0]) = 1 then
            begin
              if findpart('POLYGON', j) then
              begin
                s := format('Create Region into Variable ob %d ', [num_polygons(j+1)]);
                s := s + points(sl2[j]);
                for j := j+1 to sl2.Count-1 do
                  if sl2[j][1] = '(' then
                    s := s + points(sl2[j])
                  else
                    break;
                miDo(s);
                miDo('Insert into ErrorTable (obj,КН,_errors) Values(ob, "%s", "%s")',
                  [cn, 'Часть объекта']);
              end;

              if findpart('LINESTRING', j) then
              begin
                s := format('Create pline into Variable ob multiple %d ', [num_polygons(j+1)]);
                s := s + points(sl2[j]);
                for j := j+1 to sl2.Count-1 do
                  if sl2[j][1] = '(' then
                    s := s + points(sl2[j])
                  else
                    break;
                miDo(s);
                miDo('Insert into ErrorTable (obj,КН,_errors) Values(ob, "%s", "%s")',
                  [cn, 'Часть объекта']);
              end;

            end;
          end;

        end;
      end;

      Info('');
      if miEvalInt('TableInfo(ErrorTable, %d)', [TAB_INFO_NCOLS]) > 0 then
      begin
        Result := false;
        miDo('Commit table ErrorTable');
        miDo('Map From %s, ErrorTable', [tab]);
        miDo('Browse * From ErrorTable');
      end else
        miDo('Close table ErrorTable');
    end;

  finally
    sl0.Free;
    sl.Free;
    sl2.Free;
    ini.Free;
    FreeNil(ds);
  end;
end;

end.
