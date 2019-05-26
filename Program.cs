using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;

using System.IO;
using System.IO.BACnet;

using System.IO.BACnet.Serialize;
using System.IO.BACnet.Storage;

namespace AHUScheduleEvaluator
{
  class Program
  {

    #region 定数宣言

    private const uint BUILDING_NUMBER = 1;

    private static readonly char[] DBL_Q = new char[] { '"' };

    // define error code
    private const int ERROR_BACNET_SERVICE_TIMEOUT = -1;
    private const int ERROR_BACNET_ADDRESS_LIST_NOT_FOUND = -2;
    private const int ERROR_AHU_SETTING_NOT_FOUND = -3;
    private const int ERROR_UNHANDLED_EXCEPTION = -100;

    #endregion

    #region static変数

    /// <summary>BACnetアドレスリスト</summary>
    private static Dictionary<uint, BacnetAddress> addList = new Dictionary<uint, BacnetAddress>();
    /// <summary>実行ファイルのディレクトリ</summary>
    private static string dir = "";
    /// <summary>起動する外部エミュレータのプロセス</summary>
    private static Process emulator = null;
    #endregion

    #region 列挙型定義

    #region 一般

    /// <summary>Object Indices</summary>
    internal enum obj_Indices_General
    {
      /// <summary>日時</summary>
      obj_dateTime,
      /// <summary>外気温度[C]</summary>
      obj_outdoorTemp,
      /// <summary>外気湿度[%]</summary>
      obj_outdoorHumd,
      /// <summary>季節1</summary>
      obj_calendar1,
      /// <summary>季節2</summary>
      obj_calendar2,
      /// <summary>季節3</summary>
      obj_calendar3,
      /// <summary>季節4</summary>
      obj_calendar4,
      /// <summary>季節5</summary>
      obj_calendar5,
      /// <summary>季節6</summary>
      obj_calendar6,
      /// <summary>季節7</summary>
      obj_calendar7,
      /// <summary>季節8</summary>
      obj_calendar8,
      /// <summary>季節9</summary>
      obj_calendar9,
      /// <summary>季節10</summary>
      obj_calendar10,
      /// <summary>季節11</summary>
      obj_calendar11,
      /// <summary>季節12</summary>
      obj_calendar12,
      /// <summary>季節13</summary>
      obj_calendar13,
      /// <summary>季節14</summary>
      obj_calendar14,
      /// <summary>季節15</summary>
      obj_calendar15,
      /// <summary>季節16</summary>
      obj_calendar16,
      /// <summary>季節17</summary>
      obj_calendar17,
      /// <summary>季節18</summary>
      obj_calendar18,
      /// <summary>季節19</summary>
      obj_calendar19,
      /// <summary>季節20</summary>
      obj_calendar20,
      /// <summary>季節1名称</summary>
      obj_calendar1_name,
      /// <summary>季節2名称</summary>
      obj_calendar2_name,
      /// <summary>季節3名称</summary>
      obj_calendar3_name,
      /// <summary>季節4名称</summary>
      obj_calendar4_name,
      /// <summary>季節5名称</summary>
      obj_calendar5_name,
      /// <summary>季節6名称</summary>
      obj_calendar6_name,
      /// <summary>季節7名称</summary>
      obj_calendar7_name,
      /// <summary>季節8名称</summary>
      obj_calendar8_name,
      /// <summary>季節9名称</summary>
      obj_calendar9_name,
      /// <summary>季節10名称</summary>
      obj_calendar10_name,
      /// <summary>季節11名称</summary>
      obj_calendar11_name,
      /// <summary>季節12名称</summary>
      obj_calendar12_name,
      /// <summary>季節13名称</summary>
      obj_calendar13_name,
      /// <summary>季節14名称</summary>
      obj_calendar14_name,
      /// <summary>季節15名称</summary>
      obj_calendar15_name,
      /// <summary>季節16名称</summary>
      obj_calendar16_name,
      /// <summary>季節17名称</summary>
      obj_calendar17_name,
      /// <summary>季節18名称</summary>
      obj_calendar18_name,
      /// <summary>季節19名称</summary>
      obj_calendar19_name,
      /// <summary>季節20名称</summary>
      obj_calendar20_name,
      /// <summary>太陽高度[radian]</summary>
      obj_sunAlt,
      /// <summary>外気方位角[radian]</summary>
      obj_sunOri,
    }

    #endregion

    #endregion

    #region メイン処理

    static int Main(string[] args)
    {

      //未処理例外をキャッチするイベントハンドラ
      System.Threading.Thread.GetDomain().UnhandledException += new
        UnhandledExceptionEventHandler(UnhandledException);

      //実行ファイルのディレクトリを取得
      dir = System.Reflection.Assembly.GetExecutingAssembly().Location;
      dir = dir.Remove(dir.LastIndexOf("\\"));

      // アドレスリストの最終更新時刻を取得する
      DateTime lastUpdated = File.GetLastWriteTime(dir + Path.DirectorySeparatorChar + "ExclusivePort.csv");

      // エミュレータを起動
      Console.WriteLine("Starting Emulatior.");
      emulator = Process.Start(dir + Path.DirectorySeparatorChar + "Shizuku.exe");

      // BACnet通信が確立してアドレスリストが更新されるのを待つ
      Console.Write("Waiting BACnet service.");
      int count = 0;
      while (File.GetLastWriteTime(dir + Path.DirectorySeparatorChar + "ExclusivePort.csv") == lastUpdated)
      {
        System.Threading.Thread.Sleep(1000);
        Console.Write(".");
        if (count++ > 300)
        {
          WriteErrorMessage("Timeout BACnet service starting process.");
          Shutdown();
          return ERROR_BACNET_SERVICE_TIMEOUT;
        }
      }
      Console.WriteLine("\nBACnet service started.");

      //BACnet Deviceのアドレスリストを取得する
      Console.WriteLine("Getting BACnet address list.");
      if (!File.Exists(dir + Path.DirectorySeparatorChar + "ExclusivePort.csv"))
      {
        WriteErrorMessage("Couldn't find \"ExclusivePort.csv\".");
        Shutdown();
        return ERROR_BACNET_ADDRESS_LIST_NOT_FOUND;
      }
      using (StreamReader streamReader = new StreamReader(dir + Path.DirectorySeparatorChar + "ExclusivePort.csv"))
      {
        string ipAdd = streamReader.ReadLine();
        string bf1;
        while ((bf1 = streamReader.ReadLine()) != null)
        {
          string[] bf2 = bf1.Split(',');
          addList.Add
            (uint.Parse(bf2[0]), new BacnetAddress(BacnetAddressTypes.IP, ipAdd + ":" + bf2[1]));
        }
      }

      // 設定データを読み込む
      Console.WriteLine("Get AHU setting data.");
      if (!File.Exists(dir + Path.DirectorySeparatorChar + "settemp.csv") || !File.Exists(dir + Path.DirectorySeparatorChar + "settime.csv"))
      {
        WriteErrorMessage("Couldn't find setpoint.csv and/or setstartstop.csv.");
        Shutdown();
        return ERROR_AHU_SETTING_NOT_FOUND;
      }
      double[][] setTempArray = ReadCsv(dir + Path.DirectorySeparatorChar + "settemp.csv");
      double[][] setTimeArray = ReadCsv(dir + Path.DirectorySeparatorChar + "settime.csv");

      // 起動停止時刻と温度設定の書き込み
      ScheduleAHU(setTimeArray, setTempArray);

      // 時刻が次の日になるまで待つ
      Console.WriteLine("Waiting 8/22");
      WaitEmulation(new DateTime(2019, 8, 22, 0, 0, 0));

      // シャットダウン
      Shutdown();

      // プログラムの終了
      return 0;
    }

    public static void Shutdown()
    {
      try
      {
        // プロセスが落ちるまでkillを繰り返し
        do
        {
          emulator.Kill();
        } while (IsProcessRunning("shizuku.exe", emulator.Id));
      }
      catch (Exception e)
      {
        WriteErrorMessage("Exception occured in shutdown()");
        WriteErrorMessage(e.Message);
      }
      finally
      {
        emulator.Close();
        emulator.Dispose();
      }

    }

    public static void UnhandledException(object sender, UnhandledExceptionEventArgs e)
    {
      Exception ex = e.ExceptionObject as Exception;
      if (ex != null)
      {
        // 例外のメッセージ出力
        string message = "Exception notification of UnhandledException in AHUScheduleEvaluator. \n" +
        "  error: " + ex.Message + "\n" + "  stack trace: \n" + ex.StackTrace;
        WriteErrorMessage(message);
      }

      //エミュレータ落とす
      Shutdown();

      //プログラムの終了
      Environment.Exit(ERROR_UNHANDLED_EXCEPTION);
    }

    public static void WriteErrorMessage(string message)
    {
      // 上書きでerror.logに出力
      using (StreamWriter streamWriter = new StreamWriter(dir + Path.DirectorySeparatorChar + "error.log", true))
      {
        streamWriter.Write("\nTime: " + DateTime.Now.ToString() + " ");
        streamWriter.Write("Program directory: " + dir + "\n");
        streamWriter.Write("  " + message + "\n");
      }
    }

    // 指定されたID，名称のプロセスが生きているか確認する．
    public static bool IsProcessRunning(string name, int id)
    {
      System.Diagnostics.Process[] processArray = System.Diagnostics.Process.GetProcessesByName(name);
      foreach (System.Diagnostics.Process process in processArray)
      {
        if (id == process.Id) return true;
        else return false;
      }
      return false;
    }

    public static void WaitEmulation(DateTime end)
    {
      //BACnetクライアント作成
      BacnetClient client = new BacnetClient(new BacnetIpUdpProtocolTransport(0xBAC0, false));
      client.WritePriority = 7; // 優先度を7にして他の制御より上書きする
      client.Start();

      int count = 0;

      // エミュレータ時刻がendを超過するのを待つ
      DateTime now = DateTime.Now;
      do{
        System.Threading.Thread.Sleep(1000);  //1秒待ち
        if (getDateTime(client, addList, out now))
        {
          Console.WriteLine(now.ToString());
        }
        else
        {
          Console.WriteLine("ReadDateTime error.");
          throw new Exception("Error in read DateTime.");
        }
        // タイマカウントで600秒経過で終了
        if (count++ > 600)
        {
          throw new Exception("Timeout emulation.");
        }
      } while (end.CompareTo(now) == 1) ;

      client.Dispose();
    }

    public static double[][] ReadCsv(string fileName)
    {
      List<double[]> readCsvList = new List<double[]>();
      using (StreamReader readCsvObject = new StreamReader(fileName, Encoding.GetEncoding("utf-8")))
      {
        readCsvObject.ReadLine(); // 1行目は読み飛ばす
        while (!readCsvObject.EndOfStream)
        {
          string readCsvLine = readCsvObject.ReadLine();
          string[] lineElements = readCsvLine.Split(',');
          // 末尾か空文字なら削除
          if ( lineElements[lineElements.Length-1].Equals("") ) {
            List<string> temp = new List<string>();
            for (int i=0; i < lineElements.Length - 1; i++) temp.Add(lineElements[i]);
            lineElements = temp.ToArray();
          }
          double[] casted = Array.ConvertAll(lineElements, double.Parse);
          readCsvList.Add(casted);
        }
      }
      return readCsvList.ToArray();
    }
    #endregion

    #region その他の処理

    private static void ScheduleAHU(double[][] setTimeArray, double[][] setTempArray)
    {
      //BACnetクライアント作成
      BacnetClient client = new BacnetClient(new BacnetIpUdpProtocolTransport(0xBAC0, false));
      client.WritePriority = 7; // 優先度を7にして他の制御より上書きする
      client.Start();

      //カレンダ設定***************
      int scheduleNumber = 1;
      Console.WriteLine("Set Calendar Number of Schedule " + (scheduleNumber + 1));
      for (int ahu = 0; ahu < setTimeArray.Length; ahu++)
      {
        uint ahuID = makeAHUDeviceID((uint)setTimeArray[ahu][0], (uint)setTimeArray[ahu][1]);

        Console.Write("Setting " + ahuID + "...  ");
        IList<BacnetValue> vals = new List<BacnetValue>();
        if (client.WritePropertyRequest(
          addList[ahuID],
          new BacnetObjectId(BacnetObjectTypes.OBJECT_ANALOG_OUTPUT, (uint)scheduleNumber),
          BacnetPropertyIds.PROP_PRESENT_VALUE,
          new BacnetValue[] { new BacnetValue(scheduleNumber + 2) }))
          Console.WriteLine("Success");
        else Console.WriteLine("Failed");
      }
      Console.WriteLine();

      //時刻別運転モード設定****************
      //月曜日～金曜日までsetTimeのスケジュールにし，土日は常時ShutOffとする

      Console.WriteLine("Setting Operating Mode Schedule " + (scheduleNumber + 1));
      
      //AHUごとに運転開始時刻と運転停止時刻を設定
      for (int ahu = 0; ahu < setTimeArray.Length; ahu++)
      {
        // Manual ASN.1/BER encoding buffer初期化
        EncodeBuffer buffer = client.GetEncodeBuffer(0);
        ASN1.encode_opening_tag(buffer, 3);

        //開始時刻
        uint startHour = (uint)Math.Floor(setTimeArray[ahu][2]);
        uint startMinute = (uint)Math.Floor((setTimeArray[ahu][2] - (double)startHour) * 60);
        //運転終了時刻
        uint endHour = (uint)Math.Floor(setTimeArray[ahu][2] + setTimeArray[ahu][3]);
        uint endMinute = (uint)Math.Floor((setTimeArray[ahu][2] + setTimeArray[ahu][3] - (double)endHour) * 60);

        // 開始時刻と運転時刻をバッファに蓄積
        string cooling = "2";
        string mode = cooling;
        string shutoff = "0";

        for (int day = 0; day < 7; day++)
        {
          // 1日の開始
          ASN1.encode_opening_tag(buffer, 0);

          // 0:00に停止
          BacnetValue bdtInit = Property.DeserializeValue("00:00:00", BacnetApplicationTags.BACNET_APPLICATION_TAG_TIME);
          BacnetValue bvalInit = Property.DeserializeValue(shutoff, BacnetApplicationTags.BACNET_APPLICATION_TAG_SIGNED_INT);
          ASN1.bacapp_encode_application_data(buffer, bdtInit);
          ASN1.bacapp_encode_application_data(buffer, bvalInit);

          // 土日は空調機起動しない
          if (day < 5) {
            // 開始時刻に開始
            BacnetValue bdtStart = Property.DeserializeValue(startHour.ToString() + ":" + startMinute.ToString() + ":00", BacnetApplicationTags.BACNET_APPLICATION_TAG_TIME);
            BacnetValue bvalStart = Property.DeserializeValue(mode, BacnetApplicationTags.BACNET_APPLICATION_TAG_SIGNED_INT);
            ASN1.bacapp_encode_application_data(buffer, bdtStart);
            ASN1.bacapp_encode_application_data(buffer, bvalStart);

            // 停止時刻に停止
            BacnetValue bdtStop = Property.DeserializeValue(endHour.ToString() + ":" + endMinute.ToString() + ":00", BacnetApplicationTags.BACNET_APPLICATION_TAG_TIME);
            BacnetValue bvalStop = Property.DeserializeValue(shutoff, BacnetApplicationTags.BACNET_APPLICATION_TAG_SIGNED_INT);
            ASN1.bacapp_encode_application_data(buffer, bdtStop);
            ASN1.bacapp_encode_application_data(buffer, bvalStop);
          }
          // 1日の終わり
          ASN1.encode_closing_tag(buffer, 0);
        }
        // 締め
        ASN1.encode_closing_tag(buffer, 3);
        Array.Resize<byte>(ref buffer.buffer, buffer.offset);

        // AHUのデバイスIDを作成
        uint ahuID = makeAHUDeviceID((uint)setTimeArray[ahu][0], (uint)setTimeArray[ahu][1]);

        // 起動停止スケジュールの書き込み
        byte[] InOutBuffer = buffer.buffer;
        Console.Write("Setting " + ahuID + "...  ");
        if (client.RawEncodedDecodedPropertyConfirmedRequest(
          addList[ahuID],
          new BacnetObjectId(BacnetObjectTypes.OBJECT_SCHEDULE, (uint)(scheduleNumber + 4)),
          BacnetPropertyIds.PROP_WEEKLY_SCHEDULE,
          BacnetConfirmedServices.SERVICE_CONFIRMED_WRITE_PROPERTY,
          ref InOutBuffer))
          Console.WriteLine("Success");
        else Console.WriteLine("Failed");
      }
      Console.WriteLine();

      // --------------------------------------------------------------------------

      //時刻別温度設定値設定****************
      // AHUの給気温度はスケジュールではなくAOのポイントで設定．
      Console.WriteLine("Setting Setpoint Temperature in each zone for season " + (scheduleNumber + 1));
      for (int index = 0; index < setTempArray.Length; index++)
      {
        // AHUのデバイスIDを計算
        uint ahuID = makeAHUDeviceID((uint)setTempArray[index][0], (uint)setTempArray[index][1]);

        // 設定する温度
        double setPoint = setTempArray[index][3];
        setPoint = Math.Max(16, Math.Min(30, setPoint));

        // インスタンスNoを計算．20+4*(zone-1)+(season-1)
        bacnetObjectID bID = new bacnetObjectID(bacnetObjectType.analogOutput, (uint)(20 + 4*(setTempArray[index][2]-1) + scheduleNumber ));

        // 書き込み実施
        Console.Write("Setting " + ahuID + "-zone "+ setTempArray[index][2] +"...  ");
        bool result = writePropertyRequest(client, addList, ahuID, bID, setPoint, 7);
        if (result) Console.WriteLine("Success");
        else Console.WriteLine("Failed");
      }

      Console.WriteLine();
    }

    // AHUのBACnetデバイスIDを階数とAHU番号から計算する
    public static uint makeAHUDeviceID(uint floor, uint ahuNumber)
    {
      uint ahuID = 104000 + floor * 100 + ahuNumber * 1;
      return ahuID;
    }

    // 列挙型定義
    protected enum bacnetObjectType
    {
      analogInput,
      analogOutput,
      binaryInput,
      binaryOutput,
      datetimeValue,
    }
    private static BacnetObjectTypes convertType(bacnetObjectType type)
    {
      switch (type)
      {
        case bacnetObjectType.analogInput:
          return BacnetObjectTypes.OBJECT_ANALOG_INPUT;
        case bacnetObjectType.analogOutput:
          return BacnetObjectTypes.OBJECT_ANALOG_OUTPUT;
        case bacnetObjectType.binaryInput:
          return BacnetObjectTypes.OBJECT_BINARY_INPUT;
        case bacnetObjectType.binaryOutput:
          return BacnetObjectTypes.OBJECT_BINARY_OUTPUT;
        case bacnetObjectType.datetimeValue:
          return BacnetObjectTypes.OBJECT_DATETIME_VALUE;
        default:
          return BacnetObjectTypes.OBJECT_ANALOG_INPUT;
      }
    }

    /// <summary>BACnetObjectを特定する情報を保持する</summary>
    protected class bacnetObjectID
    {
      /// <summary>オブジェクトのタイプ</summary>
      public bacnetObjectType type { get; private set; }

      /// <summary>インスタンス番号</summary>
      public uint instance { get; private set; }

      /// <summary>インスタンスを初期化する</summary>
      /// <param name="type">オブジェクトのタイプ</param>
      /// <param name="instance">インスタンス番号</param>
      public bacnetObjectID(bacnetObjectType type, uint instance)
      {
        this.type = type;
        this.instance = instance;
      }
    }

    /// <summary>ReadPropertyを実行する</summary>
    /// <param name="deviceID">BACnetDeviceID</param>
    /// <param name="objectID">オブジェクトID</param>
    /// <param name="presentValue">出力：現在値</param>
    /// <returns>通信成功の真偽</returns>
    protected static bool readPropertyRequest
      (BacnetClient client, Dictionary<uint, BacnetAddress> addressList, uint deviceID, bacnetObjectID objectID, out object presentValue)
    {
      presentValue = null;
      BacnetObjectId bId = new BacnetObjectId(convertType(objectID.type), objectID.instance);
      IList<BacnetValue> vals = new List<BacnetValue>();
      if (client.ReadPropertyRequest(addressList[deviceID], bId, BacnetPropertyIds.PROP_PRESENT_VALUE, out vals))
      {
        presentValue = vals[0].Value;
        return true;
      }
      else return false;
    }

    private static bool getDateTime(BacnetClient client, Dictionary<uint, BacnetAddress> addressList, out DateTime dtNow)
    {
      IList<BacnetValue> vals = new List<BacnetValue>();
      BacnetObjectId bId = new BacnetObjectId(BacnetObjectTypes.OBJECT_DATETIME_VALUE, 0);
      if (client.ReadPropertyRequest(addressList[100000], bId, BacnetPropertyIds.PROP_PRESENT_VALUE, out vals))
      {
        DateTime day = (DateTime)vals[0].Value;
        DateTime time = (DateTime)vals[1].Value;
        dtNow = new DateTime(day.Year, day.Month, day.Day, time.Hour, time.Minute, time.Second);
        return true;
      }
      else
      {
        dtNow = DateTime.Now;
        return false;
      }
    }

    /// <summary>WritePropertyを実行する</summary>
    /// <param name="deviceID">BACnetDeviceID</param>
    /// <param name="objectID">オブジェクトID</param>
    /// <param name="value">書き込む現在値</param>
    /// <param name="writePriority">書き込みのプライオリティ（7以下で標準スケジュールに優先）</param>
    /// <returns>通信成功の真偽</returns>
    protected static bool writePropertyRequest
      (BacnetClient client, Dictionary<uint, BacnetAddress> addressList, uint deviceID, bacnetObjectID objectID, object value, uint writePriority)
    {
      client.WritePriority = writePriority;

      BacnetObjectId bId = new BacnetObjectId(convertType(objectID.type), objectID.instance);
      BacnetValue[] bVals = new BacnetValue[] { new BacnetValue(value) };
      return client.WritePropertyRequest(addressList[deviceID], bId, BacnetPropertyIds.PROP_PRESENT_VALUE, bVals);
    }

    private static void writePresentValue(StreamReader sReader)
    {
      //BACnetクライアント作成
      BacnetClient client = new BacnetClient(new BacnetIpUdpProtocolTransport(0xBAC0, false));
      client.WritePriority = 7;
      client.Start();

      sReader.ReadLine(); //空行
      string buff1;
      while ((buff1 = sReader.ReadLine()) != null)
      {
        buff1 = buff1.Replace("\"", "");
        string[] buff2 = buff1.Split(',');
        uint dvID = uint.Parse(buff2[0]);
        uint instNumber = uint.Parse(buff2[1]);

        Console.WriteLine("WritePropertyRequest Device ID=" + dvID + ", Instance Number=" + instNumber + "...");
        BacnetObjectId bId;
        IList<BacnetValue> vals = new List<BacnetValue>();
        //ANALOG_OUTPUTの場合
        if (buff2[2] == "OBJECT_ANALOG_OUTPUT")
        {
          bId = new BacnetObjectId(BacnetObjectTypes.OBJECT_ANALOG_OUTPUT, instNumber);
          vals.Add(new BacnetValue(double.Parse(buff2[3])));
        }
        //BINARY_OUTPUTの場合
        else
        {
          bId = new BacnetObjectId(BacnetObjectTypes.OBJECT_BINARY_OUTPUT, instNumber);
          bool tf = (buff2[3] == "#TRUE#" || buff2[3] == "1");
          vals.Add(new BacnetValue(tf ? (uint)1 : (uint)0));
        }
        //接続
        if (client.WritePropertyRequest(addList[dvID], bId, BacnetPropertyIds.PROP_PRESENT_VALUE, vals))
          Console.WriteLine("Success");
        else Console.WriteLine("Failed");
      }
      WriteMessage("Communication end");

    }

    private static void WriteMessage(string msg)
    {
      Console.WriteLine(msg);
      Console.WriteLine("Press any key to continue.");
      Console.ReadLine();
    }

    #endregion

  }
}
