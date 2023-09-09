import {
  InboxOutlined,
  QuestionCircleTwoTone,
  SettingTwoTone,
} from "@ant-design/icons";
import type { UploadProps } from "antd";
import {
  Collapse,
  ConfigProvider,
  FloatButton,
  message,
  Modal,
  Tag,
  Upload,
} from "antd";
import { Data } from "./types";
import "./App.scss";
import { useEffect, useState } from "react";
import logo from "@/assets/logo.svg";

const { Dragger } = Upload;

function App() {
  const [allTopCounts, setAllTopCounts] = useState<number>(
    Number(localStorage.getItem("allTopCounts") || "12"),
  );
  const [allSecondCounts, setAllSecondCounts] = useState<number>(
    Number(localStorage.getItem("allSecondCounts") || "38"),
  );
  const [allThirdCounts, setAllThirdCounts] = useState<number>(
    Number(localStorage.getItem("allThirdCounts") || "127"),
  );
  const [allPoorCounts, setAllPoorCounts] = useState<number>(
    Number(localStorage.getItem("allPoorCounts") || "26"),
  );

  const [isSettingsModalOpen, setIsSettingsModalOpen] = useState(false);
  const [isHelpModalOpen, setIsHelpModalOpen] = useState(false);

  const showSettingsModal = () => {
    setIsSettingsModalOpen(true);
  };

  const handleSettingsOk = () => {
    setIsSettingsModalOpen(false);
  };

  const handleSettingsCancel = () => {
    setIsSettingsModalOpen(false);
  };

  const showHelpModal = () => {
    setIsHelpModalOpen(true);
  };

  const handleHelpOk = () => {
    setIsHelpModalOpen(false);
  };

  const handleHelpCancel = () => {
    setIsHelpModalOpen(false);
  };

  const props: UploadProps = {
    name: "file",
    multiple: true,
    // 限制格式为 .xlsx
    accept: ".xlsx",
    showUploadList: false,
    customRequest: (options) => {
      const Excel = require("exceljs");
      const { onSuccess, onError, file } = options;
      // 读取 file 并将其转化为 arrayBuffer
      const reader = new FileReader();
      reader.readAsArrayBuffer(file as Blob);
      reader.onload = async (e) => {
        // 读取 file 并将其转化为 XLSX 可以解析的类型
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = new Excel.Workbook();
        const sheet_meta = await workbook.xlsx.load(data);
        const sheet = sheet_meta._worksheets[1];

        const rowCount = sheet.rowCount; // 行数
        const columnCount = sheet.columnCount; // 列数
        const studentCount = rowCount - 1; // 参与学生数

        // 创建一个新的 xlsx 文件，并创建一个新的 sheet
        const newWorkbook = new Excel.Workbook();
        const newSheet = newWorkbook.addWorksheet("Sheet1");

        let nowTopCounts = 0;
        let nowSecondCounts = 0;
        let nowThirdCounts = 0;
        let nowPoorCounts = 0;

        // 将原来的属性值复制到新的单元格
        for (let i = 1; i <= rowCount; i++) {
          // 获取第 i 行
          const row = sheet.getRow(i);
          // 创建一个数组，值从 1 到 columnCount + 1
          const arr = Array.from({ length: columnCount + 1 }, (_, i) => i + 1);
          arr.map((j) => {
            newSheet.getCell(`${String.fromCharCode(64 + j)}${i}`).value =
              row.getCell(j).text;
            if (j === columnCount + 1 && i === 1) {
              newSheet.getCell(`${String.fromCharCode(64 + j)}${i}`).value =
                "奖学金1";
            }
            // 居中，添加边框
            newSheet.getCell(`${String.fromCharCode(64 + j)}${i}`).alignment = {
              vertical: "middle",
              horizontal: "center",
            };
            newSheet.getCell(`${String.fromCharCode(64 + j)}${i}`).font = {
              name: "等线",
            };
            newSheet.getCell(`${String.fromCharCode(64 + j)}${i}`).border = {
              top: { style: "thin" },
              left: { style: "thin" },
              bottom: { style: "thin" },
              right: { style: "thin" },
            };
            // 对于第一行
            if (i === 1) {
              // 将第一行设置为可以换行
              newSheet.getCell(`${String.fromCharCode(64 + j)}${i}`).alignment =
                {
                  wrapText: true,
                  vertical: "middle",
                  horizontal: "center",
                };
              newSheet.getCell(`${String.fromCharCode(64 + j)}${i}`).font = {
                name: "等线",
                bold: true,
              };
              // 将第一行固定在第一行
              newSheet.views = [
                {
                  state: "frozen",
                  xSplit: 0,
                  ySplit: 1,
                },
              ];
            }
            // 设置单元格的宽度
            newSheet.getColumn(j).width = 15;
          });
        }

        // 学生数的前 10%，向上取整
        const studentCount10 = Math.ceil(studentCount * 0.1);
        // 学生数的前 15%
        const studentCount15 = Math.ceil(studentCount * 0.15);
        // 学生数的前 33%
        const studentCount33 = Math.ceil(studentCount * 0.33);
        // 学生数的前 40%
        const studentCount40 = Math.ceil(studentCount * 0.4);

        function jurgeTop(data: Data) {
          if (
            data.examRank <= studentCount10 &&
            data.combinRank <= studentCount10 &&
            data.examScore >= 80
          ) {
            nowTopCounts++;
            if (nowTopCounts > allTopCounts) {
              return false;
            }
            return true;
          }
          return false;
        }

        function jurgeSecond(data: Data) {
          if (
            data.examRank <= studentCount15 &&
            data.combinRank <= studentCount15 &&
            data.examScore >= 78
          ) {
            nowSecondCounts++;
            if (nowSecondCounts > allSecondCounts) {
              return false;
            }
            return true;
          }
          return false;
        }

        function jurgeThird(data: Data) {
          if (
            data.examRank <= studentCount40 &&
            data.combinRank <= studentCount40 &&
            data.examScore >= 75
          ) {
            nowThirdCounts++;
            if (nowThirdCounts > allThirdCounts) {
              return false;
            }
            return true;
          }
          return false;
        }

        function jurgeBase(data: Data) {
          if (
            data.volunteetTime >= 20 &&
            data.apartmentPoints >= 0 &&
            data.passCourse !== "否" &&
            data.totalCredit >= 20 &&
            data.workHours !== "否"
          ) {
            return true;
          }
          return false;
        }

        function jurge(data: Data) {
          if (!jurgeBase(data)) return null;
          if (jurgeTop(data)) return "学业一等奖学金";
          if (jurgeSecond(data)) return "学业二等奖学金";
          if (jurgeThird(data)) return "学业三等奖学金";
          return null;
        }

        for (let i = 2; i <= rowCount; i++) {
          // 获取第 i 行
          const row = newSheet.getRow(i);
          const data = {
            examRank: Number(row.getCell(8).text),
            combinRank: Number(row.getCell(6).text),
            volunteetTime: Number(row.getCell(11).text),
            apartmentPoints: Number(row.getCell(12).text),
            passCourse: String(row.getCell(13).text),
            totalCredit: Number(Number(row.getCell(14).text)),
            workHours: String(row.getCell(15).text),
            examScore: Number(row.getCell(7).text),
            combinScore: Number(row.getCell(5).text),
            isPoor: String(row.getCell(10).text),
          };
          row.getCell(columnCount + 1).border = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" },
          };

          // 如果不满足奖学金条件，跳过
          const scholarship = jurge(data);
          const poorship = () => {
            if (
              data.isPoor === "是" &&
              data.combinRank <= studentCount33 &&
              data.examRank <= studentCount33
            ) {
              nowPoorCounts++;
              if (nowPoorCounts > allPoorCounts) {
                return false;
              }
              return true;
            }
            return false;
          };
          if (poorship()) {
            row.getCell(columnCount + 2).value = "国家励志奖学金";
          }
          if (!scholarship) continue;
          if (scholarship) {
            // 在第 i 行的最后一列添加奖学金类型
            row.getCell(columnCount + 1).value = scholarship;
            const color =
              scholarship === "学业一等奖学金"
                ? "f8cbad"
                : scholarship === "学业二等奖学金"
                ? "bdd7ee"
                : scholarship === "学业三等奖学金"
                ? "ffff00"
                : "FFFFFFFF";
            // 创建一个数组，值从 1 到 columnCount + 1
            const arr = Array.from(
              { length: columnCount + 1 },
              (_, i) => i + 1,
            );
            arr.map((j) => {
              row.getCell(j).style.fill = {
                type: "pattern",
                pattern: "solid",
                fgColor: { argb: color },
              };
              row.getCell(j).protection = {
                locked: false,
              };
            });
          }
        }
        await newWorkbook.xlsx.writeFile(`${(file as any).name}-result.xlsx`);
        onSuccess!("done");
      };
      // 如果加载失败
      reader.onerror = () => {
        onError!(new Error("读取文件失败"));
      };
    },
    onChange(info) {
      const { status } = info.file;
      if (status !== "uploading") {
        // console.log(info.file, info.fileList);
      }
      if (status === "done") {
        message.success(`${info.file.name} 分析成功`);
      } else if (status === "error") {
        message.error(`${info.file.name} 加载失败`);
      }
    },
  };

  useEffect(() => {
    console.log(allTopCounts, allSecondCounts, allThirdCounts);
    localStorage.setItem("allTopCounts", JSON.stringify(allTopCounts));
    localStorage.setItem("allSecondCounts", JSON.stringify(allSecondCounts));
    localStorage.setItem("allThirdCounts", JSON.stringify(allThirdCounts));
  }, [allTopCounts, allSecondCounts, allThirdCounts]);

  return (
    <ConfigProvider
      theme={{
        token: {
          colorPrimary: "#302d4c",
        },
      }}
    >
      <div className="App">
        <div
          style={{
            display: "flex",
            justifyContent: "center",
            alignItems: "center",
            height: 200,
            paddingTop: 200,
            paddingBottom: 0,
            overflowY: "hidden",
          }}
        >
          <img
            style={{
              width: 512,
              height: 512,
              // 不可选中
              userSelect: "none",
            }}
            src={logo}
            alt=""
          />
        </div>
        <div
          style={{
            display: "flex",
            justifyContent: "center",
            alignItems: "center",
            marginBottom: 50,
            marginTop: -50,
            flexDirection: "column",
          }}
        >
          <div
            style={{
              color: "white",
            }}
          >
            使用前先查看右下角的<strong>帮助</strong>和<strong>设置</strong>
          </div>
          <Dragger {...props} className="download-box">
            <p className="ant-upload-drag-icon">
              <InboxOutlined className="inner-down" />
            </p>
            <p className="ant-upload-text">点击或者拖动上传 .xlsx 文件</p>
            <p className="ant-upload-hint">支持单个或多个 .xlsx 文件上传</p>
          </Dragger>
        </div>
      </div>
      <FloatButton
        icon={<QuestionCircleTwoTone twoToneColor="#302d4c" />}
        onClick={() => showHelpModal()}
      />
      <Modal
        title="帮助"
        open={isHelpModalOpen}
        onOk={handleHelpOk}
        onCancel={handleHelpCancel}
        cancelText="返回"
        okText="确认"
      >
        <div
          style={{
            lineHeight: "1.8rem",
          }}
        >
          1. 表格内属性必须按照 <Tag>序号</Tag> <Tag>学号</Tag> <Tag>姓名</Tag>
          <Tag>班级</Tag>
          <Tag>综合素质测评成绩</Tag> <Tag>综合素质测评成绩排名</Tag>
          <Tag>智育测评分</Tag> <Tag>智育测评分排名</Tag> <Tag>扣分</Tag>
          <Tag>是否家庭经济困难学生</Tag> <Tag>志愿时长</Tag>{" "}
          <Tag>公寓积分总分</Tag>
          <Tag>课程是否一次性通过</Tag> <Tag>总学分</Tag>{" "}
          <Tag>劳动时长是否合格</Tag>的<strong>顺序</strong>
          写入，属性名可以改变。
        </div>
        <div
          style={{
            lineHeight: "1.8rem",
          }}
        >
          2. 必须将需要分析的表单放在 .xlsx 文件的<strong>第一个表单</strong>
          ，否则会出现错误。
        </div>
        <div
          style={{
            lineHeight: "1.8rem",
          }}
        >
          3. 分析后的文件将会直接输出在<strong>该应用所在文件夹内</strong>
        </div>
      </Modal>
      <FloatButton
        style={{
          right: 80,
        }}
        icon={<SettingTwoTone twoToneColor="#302d4c" />}
        onClick={() => showSettingsModal()}
      />
      <Modal
        title="设置"
        open={isSettingsModalOpen}
        onOk={handleSettingsOk}
        onCancel={handleSettingsCancel}
        okText="返回"
        footer={null}
      >
        <div
          style={{
            color: "#302d4c",
            margin: "10px",
          }}
        >
          实时保存，无需手动保存
        </div>
        <Collapse>
          <Collapse.Panel header="学业一等奖学金人数" key="1">
            <input
              type="number"
              value={allTopCounts}
              onChange={(e) => setAllTopCounts(Number(e.target.value))}
            />
          </Collapse.Panel>
          <Collapse.Panel header="学业二等奖学金人数" key="2">
            <input
              type="number"
              value={allSecondCounts}
              onChange={(e) => setAllSecondCounts(Number(e.target.value))}
            />
          </Collapse.Panel>
          <Collapse.Panel header="学业三等奖学金人数" key="3">
            <input
              type="number"
              value={allThirdCounts}
              onChange={(e) => setAllThirdCounts(Number(e.target.value))}
            />
          </Collapse.Panel>
        </Collapse>
      </Modal>
    </ConfigProvider>
  );
}

export default App;
