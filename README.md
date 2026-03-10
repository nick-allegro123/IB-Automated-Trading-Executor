# Multi-Broker Automated Trading Executor
## 多券商自動化交易下單系統 (Python)

## 專案概述 (Project Overview)
本專案為一套針對多券商開發的自動化下單中台系統。核心目標在於解決主流交易平台（如 MultiCharts）在 API 串接上的限制，提供更靈活、高效率的執行環境。

This project is a multi-broker automated trading middleware designed to overcome API limitations in mainstream platforms, providing a flexible and high-efficiency execution environment.

## 核心技術亮點 (Key Technical Features)

## 1. 低延遲執行架構 (Ultra-Low Latency Architecture)
* **RAMDisk 訊號監測：** 透過監控 RAMDisk 虛擬磁碟實現極低延遲的訊號讀取，大幅縮短 I/O 傳輸路徑。
  * *Monitors RAMDisk for signal files to achieve ultra-low latency execution by minimizing disk I/O overhead.*
* **多執行緒處理 (Multi-threading)：** 實測可在 1 秒內處理超過 50 筆併發指令，確保市場波動時的成交速度。
  * *Capable of processing 50+ concurrent orders per second to ensure execution speed during high volatility.*

## 2. 系統穩定性與兼容性 (Stability & Compatibility)
* **高可靠性：** 支持 365 天長效運行，優化記憶體管理以實現低資源消耗。
  * *Supports 24/7 continuous operation with optimized memory management.*
* **跨平台整合：** 支援 IB, OANDA, Binance, Charles Schwab 等多家券商。(附件提供IB、OANDA兩個版本)
  * *Supports multiple brokers including Interactive Brokers, OANDA, Binance, etc.(Source code for IB and OANDA modules included)*

## 📁 檔案版本說明 (Edition Differences)
* **Developer Sandbox (開發人員版本)：** 移除驗證邏輯，專為開發人員內測設計，提升調試效率。
  * *Stripped of auth logic for rapid internal testing and development.*
* **Production (會員系統版本)：** 包含完整的身分驗證系統與授權協議。
  * *Includes full authentication and licensing modules.*

## 📁 檔案架構 (Project Structure)
* `IB_Trading_Bot.py` : 串接 Interactive Brokers TWS/Gateway 之執行模組。
  * *Execution module for Interactive Brokers TWS/Gateway.*
* `OANDA_Trading_Bot.py` : 串接 OANDA v20 REST API 之執行模組。
  * *Execution module for OANDA v20 REST API.*

## 開發動機 (Development Motivation)
早期在使用 MultiCharts 或 TradingView 時，常遇到 API 限制或平台閃退問題。為了完全掌控交易流程並規避技術風險，我決定自學開發此系統。

Developed to solve API restrictions and platform stability issues in third-party software, ensuring full control over the execution workflow.

---
## ⚠️ 免責聲明 (Disclaimer)
本專案僅供技術展示使用，不構成任何投資建議。
*For technical demonstration only. Not investment advice.*
