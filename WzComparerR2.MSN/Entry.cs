using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using System.Xml;
using DevComponents.DotNetBar;
using WzComparerR2.CharaSim;
using WzComparerR2.CharaSimControl;
using WzComparerR2.Common;
using WzComparerR2.Config;
using WzComparerR2.PluginBase;

namespace WzComparerR2.MSN
{
    public class Entry : PluginEntry
    {
        private SuperTabControlPanel panel;
        private Label infoLabel;
        private FlowLayoutPanel charaSimToolbar;
        private Button btnShowItem;
        private Button btnShowStat;
        private Button btnShowEquip;
        private Button btnShowQuick;
        private Button btnShowAll;
        private Button btnHideAll;
        private ButtonItem languageMenu;
        private ButtonItem languageEnglish;
        private ButtonItem languageChinese;
        private string currentLanguage = "en";
        private bool autoLoadedBaseWz;
        private bool autoLoadScheduled;
        private EventHandler autoLoadIdleHandler;
        private CharaSimControlGroup charaSimCtrl;

        public Entry(PluginContext context)
            : base(context)
        {
        }

        protected override void OnLoad()
        {
            Context.MainForm.Shown += MainForm_Shown;
            Context.MainForm.FormClosing += MainForm_FormClosing;
            Context.WzOpened += Context_WzOpened;
            panel = new SuperTabControlPanel();
            panel.Dock = DockStyle.Fill;
            infoLabel = new Label
            {
                AutoSize = true,
                Dock = DockStyle.Top,
                Text = "Language switch is available in the ribbon."
            };
            charaSimToolbar = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 36,
                AutoSize = false,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false
            };
            btnShowItem = CreateCharaSimButton("Item");
            btnShowStat = CreateCharaSimButton("Stat");
            btnShowEquip = CreateCharaSimButton("Equip");
            btnShowQuick = CreateCharaSimButton("QuickView");
            btnShowAll = CreateCharaSimButton("Show All");
            btnHideAll = CreateCharaSimButton("Hide All");
            btnShowItem.Click += (s, e) => ShowCharaSimItem();
            btnShowStat.Click += (s, e) => ShowCharaSimStat();
            btnShowEquip.Click += (s, e) => ShowCharaSimEquip();
            btnShowQuick.Click += (s, e) => ShowCharaSimQuickView();
            btnShowAll.Click += (s, e) => ShowCharaSimAll();
            btnHideAll.Click += (s, e) => HideCharaSimForms();
            charaSimToolbar.Controls.Add(btnShowItem);
            charaSimToolbar.Controls.Add(btnShowStat);
            charaSimToolbar.Controls.Add(btnShowEquip);
            charaSimToolbar.Controls.Add(btnShowQuick);
            charaSimToolbar.Controls.Add(btnShowAll);
            charaSimToolbar.Controls.Add(btnHideAll);
            panel.Controls.Add(charaSimToolbar);
            panel.Controls.Add(infoLabel);
            Context.AddTab("MSN", panel);
            AddLanguageMenu();
            SetMainTitle();
            ApplyLanguage();
        }

        private void MainForm_Shown(object sender, EventArgs e)
        {
            AddLanguageMenu();
            ApplyLanguage();
            ScheduleAutoLoadBaseWz();
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            HideCharaSimForms();
        }

        private void Context_WzOpened(object sender, WzStructureEventArgs e)
        {
            var form = Context.MainForm;
            if (form == null)
            {
                return;
            }
            form.BeginInvoke(new Action(EnsureCharaSimInitialized));
        }

        private void AddLanguageMenu()
        {
            if (languageMenu != null)
            {
                return;
            }

            var ribbonControl = FindControl<RibbonControl>(Context.MainForm, "ribbonControl1");
            if (ribbonControl == null)
            {
                return;
            }

            languageEnglish = new ButtonItem
            {
                Name = "msnLanguageEnglish",
                Text = "English",
                OptionGroup = "MSNLanguage",
                Checked = true
            };
            languageEnglish.Click += (s, e) => SetLanguage("en");

            languageChinese = new ButtonItem
            {
                Name = "msnLanguageChinese",
                Text = "中文",
                OptionGroup = "MSNLanguage"
            };
            languageChinese.Click += (s, e) => SetLanguage("zh");

            languageMenu = new ButtonItem
            {
                Name = "msnLanguageMenu",
                Text = "Language",
                AutoExpandOnClick = true,
                ItemAlignment = eItemAlignment.Far
            };
            languageMenu.SubItems.Add(languageEnglish);
            languageMenu.SubItems.Add(languageChinese);

            int insertIndex = ribbonControl.Items.Count;
            for (int i = 0; i < ribbonControl.Items.Count; i++)
            {
                if (string.Equals(ribbonControl.Items[i].Name, "buttonItemStyle", StringComparison.OrdinalIgnoreCase))
                {
                    insertIndex = i + 1;
                    break;
                }
            }
            ribbonControl.Items.Insert(insertIndex, languageMenu);
        }

        private static T FindControl<T>(Control root, string name) where T : Control
        {
            if (root == null)
            {
                return null;
            }
            if (string.Equals(root.Name, name, StringComparison.OrdinalIgnoreCase))
            {
                return root as T;
            }

            foreach (Control child in root.Controls)
            {
                var match = FindControl<T>(child, name);
                if (match != null)
                {
                    return match;
                }
            }

            return null;
        }

        private void SetLanguage(string language)
        {
            if (currentLanguage == language)
            {
                return;
            }

            currentLanguage = language;
            if (languageEnglish != null)
            {
                languageEnglish.Checked = language == "en";
            }
            if (languageChinese != null)
            {
                languageChinese.Checked = language == "zh";
            }
            ApplyLanguage();
        }

        private void ApplyLanguage()
        {
            var form = Context.MainForm;
            if (form == null)
            {
                return;
            }

            if (form.InvokeRequired)
            {
                form.BeginInvoke(new Action(ApplyLanguage));
                return;
            }

            bool isChinese = currentLanguage == "zh";
            ApplyMainMenuText(form, isChinese);
            languageMenu.Text = isChinese ? "语言" : "Language";
            infoLabel.Text = isChinese
                ? "使用下方按钮打开 CharaSim 控件预览。语言切换入口在右上角主题旁边。"
                : "Use the buttons below to show CharaSim controls. Language switch is next to Themes on the ribbon.";
            if (btnShowItem != null)
            {
                btnShowItem.Text = isChinese ? "物品" : "Item";
                btnShowStat.Text = isChinese ? "属性" : "Stat";
                btnShowEquip.Text = isChinese ? "装备" : "Equip";
                btnShowQuick.Text = isChinese ? "预览" : "QuickView";
                btnShowAll.Text = isChinese ? "全部显示" : "Show All";
                btnHideAll.Text = isChinese ? "全部隐藏" : "Hide All";
            }
            RefreshMainFormLayout(form);
            SetMainTitle();
        }

        private void TryAutoLoadBaseWz()
        {
            if (autoLoadedBaseWz)
            {
                return;
            }

            var form = Context.MainForm;
            if (form == null)
            {
                return;
            }

            if (HasOpenedWz(form))
            {
                autoLoadedBaseWz = true;
                return;
            }

            try
            {
                string baseWzPath = GetRecentBaseWzPath();
                if (string.IsNullOrEmpty(baseWzPath) || !File.Exists(baseWzPath))
                {
                    return;
                }

                autoLoadedBaseWz = true;
                InvokeOpenWz(form, baseWzPath);
            }
            catch (Exception ex)
            {
                PluginManager.LogError("MSN", ex, "Auto load Base.wz failed.");
            }
        }

        private void ScheduleAutoLoadBaseWz()
        {
            if (autoLoadedBaseWz || autoLoadScheduled)
            {
                return;
            }

            autoLoadScheduled = true;
            autoLoadIdleHandler = (s, e) =>
            {
                Application.Idle -= autoLoadIdleHandler;
                autoLoadIdleHandler = null;
                autoLoadScheduled = false;
                Context.MainForm.BeginInvoke(new Action(TryAutoLoadBaseWz));
            };
            Application.Idle += autoLoadIdleHandler;
        }

        private bool HasOpenedWz(Form form)
        {
            var field = form.GetType().GetField("openedWz", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            if (field == null)
            {
                return false;
            }

            if (field.GetValue(form) is System.Collections.ICollection collection)
            {
                return collection.Count > 0;
            }

            return false;
        }

        private string GetRecentBaseWzPath()
        {
            string configPath = ConfigManager.ConfigFileName;
            if (!File.Exists(configPath))
            {
                return null;
            }

            var doc = new XmlDocument();
            doc.Load(configPath);
            var nodes = doc.SelectNodes("//recentDocuments/*/@value");
            if (nodes == null)
            {
                return null;
            }

            foreach (XmlAttribute attr in nodes)
            {
                string path = attr?.Value;
                if (!string.IsNullOrEmpty(path)
                    && string.Equals(Path.GetFileName(path), "Base.wz", StringComparison.OrdinalIgnoreCase))
                {
                    return path;
                }
            }

            return null;
        }

        private void InvokeOpenWz(Form form, string wzPath)
        {
            if (form == null || string.IsNullOrEmpty(wzPath))
            {
                return;
            }

            var method = form.GetType().GetMethod("openWz", BindingFlags.Instance | BindingFlags.NonPublic);
            method?.Invoke(form, new object[] { wzPath });
        }

        private Button CreateCharaSimButton(string text)
        {
            return new Button
            {
                Text = text,
                AutoSize = true,
                Height = 28,
                Margin = new System.Windows.Forms.Padding(6, 4, 6, 4)
            };
        }

        private void EnsureCharaSimInitialized()
        {
            if (charaSimCtrl != null)
            {
                if (Context.DefaultStringLinker != null)
                {
                    charaSimCtrl.StringLinker = Context.DefaultStringLinker;
                }
                return;
            }

            charaSimCtrl = new CharaSimControlGroup
            {
                Character = new Character
                {
                    Name = "MSN"
                }
            };
            if (Context.DefaultStringLinker != null)
            {
                charaSimCtrl.StringLinker = Context.DefaultStringLinker;
            }
        }

        private void ShowCharaSimItem()
        {
            EnsureCharaSimInitialized();
            ShowCharaSimForm(charaSimCtrl.UIItem, 40, 120);
        }

        private void ShowCharaSimStat()
        {
            EnsureCharaSimInitialized();
            ShowCharaSimForm(charaSimCtrl.UIStat, 380, 120);
        }

        private void ShowCharaSimEquip()
        {
            EnsureCharaSimInitialized();
            ShowCharaSimForm(charaSimCtrl.UIEquip, 720, 120);
        }

        private void ShowCharaSimQuickView()
        {
            EnsureCharaSimInitialized();
            ShowCharaSimForm(charaSimCtrl.TooltipQuickView, 1060, 120);
        }

        private void ShowCharaSimAll()
        {
            ShowCharaSimItem();
            ShowCharaSimStat();
            ShowCharaSimEquip();
            ShowCharaSimQuickView();
        }

        private void HideCharaSimForms()
        {
            if (charaSimCtrl == null)
            {
                return;
            }

            charaSimCtrl.UIItem.Hide();
            charaSimCtrl.UIStat.Hide();
            charaSimCtrl.UIEquip.Hide();
            charaSimCtrl.TooltipQuickView.Hide();
        }

        private void ShowCharaSimForm(Form form, int offsetX, int offsetY)
        {
            if (form == null)
            {
                return;
            }

            var owner = Context.MainForm;
            if (owner == null)
            {
                return;
            }

            if (!form.Visible)
            {
                form.StartPosition = FormStartPosition.Manual;
                var origin = owner.PointToScreen(System.Drawing.Point.Empty);
                form.Location = new System.Drawing.Point(origin.X + offsetX, origin.Y + offsetY);
                form.TopMost = false;
                form.Show(owner);
            }
            form.BringToFront();
        }

        private void SetMainTitle()
        {
            var form = Context.MainForm;
            if (form == null)
            {
                return;
            }

            if (form.InvokeRequired)
            {
                form.BeginInvoke(new Action(SetMainTitle));
                return;
            }

            form.Text = "[MSN] WzComparerR2";
        }

        private void ApplyMainMenuText(Form form, bool isChinese)
        {
            var translations = GetMainMenuTranslations();
            foreach (var kvp in translations)
            {
                var targetText = isChinese ? kvp.Value.Chinese : kvp.Value.English;
                if (string.IsNullOrEmpty(targetText))
                {
                    continue;
                }

                object target = GetPrivateField(form, kvp.Key);
                if (target is BaseItem baseItem)
                {
                    baseItem.Text = targetText;
                }
                else if (target is Control control)
                {
                    control.Text = targetText;
                }
            }
        }

        private void RefreshMainFormLayout(Form form)
        {
            var ribbonControl = FindControl<RibbonControl>(form, "ribbonControl1");
            if (ribbonControl != null)
            {
                ribbonControl.RecalcLayout();
                ribbonControl.Invalidate(true);
                ribbonControl.Refresh();
            }

            form.Invalidate(true);
            form.Refresh();
        }

        private static object GetPrivateField(object instance, string name)
        {
            if (instance == null || string.IsNullOrEmpty(name))
            {
                return null;
            }

            var field = instance.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic);
            return field?.GetValue(instance);
        }

        private static Dictionary<string, Translation> GetMainMenuTranslations()
        {
            return new Dictionary<string, Translation>(StringComparer.OrdinalIgnoreCase)
            {
                { "ribbonBar8", new Translation("CharaSim", "角色模拟") },
                { "buttonItemCreateChara", new Translation("Create", "创建") },
                { "buttonItemEdit", new Translation("Edit", "编辑") },
                { "buttonItemLoadChara", new Translation("Load", "加载") },
                { "buttonItemSaveChara", new Translation("Save", "保存") },
                { "buttonItemQuickView", new Translation("Preview", "预览") },
                { "buttonItemAutoQuickView", new Translation("Auto Preview", "自动预览") },
                { "buttonItemQuickViewSetting", new Translation("Settings", "设置") },
                { "buttonItemSetItems", new Translation("Item Management", "物品管理") },
                { "buttonItemClearSetItems", new Translation("Consolidate Set Item", "合并套装物品") },
                { "buttonItemClearExclusiveEquips", new Translation("Consolidate Non-Duplicate Item", "合并非重复物品") },
                { "buttonItemClearCommodities", new Translation("Consolidate Cash Item", "合并点装物品") },
                { "buttonItemCharItem", new Translation("Inventory", "背包") },
                { "buttonItemCharaStat", new Translation("Stats", "属性") },
                { "buttonItemCharaEquip", new Translation("Equipment", "装备") },
                { "buttonItemAddItem", new Translation("Add Item", "添加物品") },
                { "ribbonBar3", new Translation("Music Player", "音乐播放器") },
                { "buttonItemLoadSound", new Translation("Load", "加载") },
                { "buttonItemSoundPlay", new Translation(" Play", "播放") },
                { "buttonItemSoundStop", new Translation("Stop", "停止") },
                { "buttonItemSoundSave", new Translation("Save", "保存") },
                { "ribbonBar9", new Translation("Patcher", "补丁") },
                { "buttonItemPatcher", new Translation("Patcher", "补丁") },
                { "ribbonBar4", new Translation("String Search", "字符串搜索") },
                { "labelItem2", new Translation("String", "字符串") },
                { "comboItem3", new Translation("All", "全部") },
                { "comboItem4", new Translation("Equipment", "装备") },
                { "comboItem5", new Translation("Item", "物品") },
                { "comboItem6", new Translation("Map", "地图") },
                { "comboItem7", new Translation("Monster", "怪物") },
                { "comboItem8", new Translation("NPC", "NPC") },
                { "comboItemSearchQuest", new Translation("Quest", "任务") },
                { "comboItem9", new Translation("Skill", "技能") },
                { "comboItem19", new Translation("Set Item", "套装") },
                { "comboItemSearchAchievement", new Translation("Achievement", "成就") },
                { "buttonItemSearchString", new Translation("Find", "查找") },
                { "buttonItemSelectStringWz", new Translation("Select Ba&se.wz", "选择 &Base.wz") },
                { "buttonItemClearStringWz", new Translation("Clear StringLinker", "清除字符串链接") },
                { "buttonItemIgnoreArticles", new Translation("Ignore Articles (a, an, the) in Search Result", "搜索结果忽略冠词 (a, an, the)") },
                { "ribbonBar1", new Translation("WZ Node Search", "WZ 节点搜索") },
                { "labelItem3", new Translation("Node", "节点") },
                { "comboItem10", new Translation("WZ Node", "WZ 节点") },
                { "comboItem11", new Translation("Image Node", "图像节点") },
                { "comboItem12", new Translation("Image Value", "图像值") },
                { "comboItem20", new Translation("Node,Value", "节点,值") },
                { "comboItem22", new Translation("Node Path", "节点路径") },
                { "buttonItemSearchWz", new Translation("Find Next", "查找下一个") },
                { "ribbonBar11", new Translation("Experimental", "实验功能") },
                { "buttonItem1", new Translation("Test (Do not use)", "测试(勿用)") },
                { "ribbonBar7", new Translation("Update", "更新") },
                { "buttonItemUpdate", new Translation("Update", "更新") },
                { "ribbonBar6", new Translation("About", "关于") },
                { "buttonItemAbout", new Translation("About", "关于") },
                { "ribbonTabItem1", new Translation("Tools", "工具") },
                { "ribbonTabItem2", new Translation("Modules", "模块") },
                { "ribbonTabItem3", new Translation("Help", "帮助") },
                { "buttonItemStyle", new Translation("Themes", "主题") },
                { "office2007StartButton1", new Translation("File", "文件") },
                { "btnItemOpenWz", new Translation("Open WZ File", "打开 WZ 文件") },
                { "btnItemOpenImg", new Translation("Open IMG", "打开 IMG") },
                { "buttonItemClose", new Translation("Close", "关闭") },
                { "buttonItemCloseAll", new Translation("Close All", "全部关闭") },
                { "labelItem8", new Translation("Recent Files", "最近文件") },
                { "btnItemOptions", new Translation("Settings", "设置") },
                { "buttonItem13", new Translation("Exit", "退出") },
                { "labelItemStatus", new Translation("Status", "状态") },
                { "buttonItemSaveImage", new Translation("Save Image", "保存图片") },
                { "buttonItemAutoSave", new Translation("Auto Save", "自动保存") },
                { "buttonItemAutoSaveFolder", new Translation("Select Folder", "选择文件夹") },
                { "buttonItemSaveWithOptions", new Translation("Custom Save", "自定义保存") },
                { "buttonItemGif", new Translation("Extract Animation", "导出动画") },
                { "buttonItemGif2", new Translation("Enable Animation Overlay", "启用动画叠加") },
                { "buttonItemExtractGifEx", new Translation("Activate+", "启用+") },
                { "buttonItemGifSetting", new Translation("Settings", "设置") },
                { "chkResolvePngLink", new Translation("Resolve Link", "解析链接") },
                { "chkEnableDarkMode", new Translation("Enable Dark Mode", "启用深色模式") },
                { "chkOutputSkillTooltip", new Translation("Save Skill Tooltip", "保存技能提示") },
                { "chkOutputCashTooltip", new Translation("Save Cash Package Tooltip", "保存点装礼包提示") },
                { "chkOutputEqpTooltip", new Translation("Save Gear Tooltip", "保存装备提示") },
                { "chkOutputItemTooltip", new Translation("Save Item Tooltip", "保存物品提示") },
                { "chkOutputMapTooltip", new Translation("Save Map Tooltip", "保存地图提示") },
                { "chkOutputMobTooltip", new Translation("Save Mob Tooltip", "保存怪物提示") },
                { "chkOutputNpcTooltip", new Translation("Save NPC Tooltip", "保存 NPC 提示") },
                { "chkOutputQuestTooltip", new Translation("Save Quest Tooltip", "保存任务提示") },
                { "chkOutputAchvTooltip", new Translation("Save Achievement Tooltip", "保存成就提示") },
                { "chkShowObjectID", new Translation("Show ID in Saved Tooltip", "提示中显示 ID") },
                { "chkShowChangeType", new Translation("Show Change Type", "显示变更类型") },
                { "chkShowLinkedTamingMob", new Translation("Show Linked Taming Mob", "显示关联坐骑") },
                { "chkSkipKMSContent", new Translation("Skip KMS Contents", "跳过 KMS 内容") },
                { "chkMseaMode", new Translation("MSEA Mode", "MSEA 模式") },
                { "chkSkipGodChangseopDuplicatedNodes", new Translation("Skip Duplicated Nodes\r\n(Ending With \"_.img\")", "跳过重复节点\r\n(以 \"_.img\" 结尾)") },
                { "chkOutputRemovedImg", new Translation("Removed Files", "已移除文件") },
                { "chkOutputAddedImg", new Translation("Added Files", "新增文件") },
                { "chkOutputPng", new Translation("PNG && Audio", "PNG 与音频") },
                { "btnEasyCompare", new Translation("Compare", "比较") },
                { "btnSkillTooltipExport", new Translation("Export Skill Tooltip", "导出技能提示") },
                { "btnExportSkillOption", new Translation("Export Skill Option", "导出技能选项") },
                { "btnExportSkill", new Translation("Export Skill", "导出技能") },
                { "btnNodeBack", new Translation("back", "后退") },
                { "btnNodeForward", new Translation("forward", "前进") },
                { "chkHashPngFileName", new Translation("Hash PNG Names", "PNG 名称哈希") },
                { "btnRootNode", new Translation("Root Node▼", "根节点▼") },
                { "btnPreset", new Translation("Presets", "预设") },
                { "btnMusicChannel", new Translation("Music Channel Owner", "音乐频道拥有者") },
                { "btnSkillChangeInfo", new Translation("Skill Change Info", "技能变更信息") },
                { "btnNewItemNews", new Translation("New Item Discoverer", "新物品发现者") },
                { "btnMapleWiki", new Translation("MapleStory Wiki Contributor", "枫之谷维基贡献者") }
            };
        }

        private sealed class Translation
        {
            public Translation(string english, string chinese)
            {
                English = english;
                Chinese = chinese;
            }

            public string English { get; }
            public string Chinese { get; }
        }

    }
}
