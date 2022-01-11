﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows;

namespace ExcelDna.IntelliSense
{
    // These are immutable representations of the state (reflecting only our interests)
    // We make a fresh a simplified state representation, so that we can make a matching state update representation.
    //
    // One shortcoming of our representation is that we don't track a second selection list and matching ExcelToolTip 
    // that might pop up for an argument, e.g. VLOOKUP's TRUE/FALSE.
    internal abstract class UIState
    {
        public static UIState ReadyState = new Ready();
        public class Ready : UIState { }
        public class FormulaEdit : UIState
        {
            public IntPtr FormulaEditWindow;    // Window where text entry focus is - either the in-cell edit window, or the formula bar
            public string FormulaPrefix;        // Never null
            public Rect EditWindowBounds;
            public IntPtr ExcelToolTipWindow;   // ExcelToolTipWindow is Zero or is _some_ visible tooltip (either for from the list or the function)

            public virtual FormulaEdit WithFormulaEditWindow(IntPtr newFormulaEditWindow)
            {
                return new FormulaEdit
                {
                    FormulaEditWindow = newFormulaEditWindow,
                    FormulaPrefix = FormulaPrefix,
                    EditWindowBounds = EditWindowBounds,
                    ExcelToolTipWindow = ExcelToolTipWindow
                };
            }

            public virtual FormulaEdit WithFormulaPrefix(string newFormulaPrefix)
            {
                return new FormulaEdit
                {
                    FormulaEditWindow = FormulaEditWindow,
                    FormulaPrefix = newFormulaPrefix ?? "",
                    EditWindowBounds = EditWindowBounds,
                    ExcelToolTipWindow = ExcelToolTipWindow
                };
            }

            public virtual FormulaEdit WithBounds(Rect newEditWindowBounds)
            {
                return new FormulaEdit
                {
                    FormulaEditWindow = FormulaEditWindow,
                    FormulaPrefix = FormulaPrefix,
                    EditWindowBounds = newEditWindowBounds,
                    ExcelToolTipWindow = ExcelToolTipWindow
                };
            }

            public virtual FormulaEdit WithToolTipWindow(IntPtr newExcelToolTipWindow)
            {
                return new FormulaEdit
                {
                    FormulaEditWindow = FormulaEditWindow,
                    FormulaPrefix = FormulaPrefix,
                    EditWindowBounds = EditWindowBounds,
                    ExcelToolTipWindow = newExcelToolTipWindow
                };
            }
        }

        public class FunctionList : FormulaEdit // CONSIDER: I'm not sure the hierarchy here has any value.
        {
            public IntPtr FunctionListWindow;
            public Rect FunctionListBounds;
            public string SelectedItemText;
            public Rect SelectedItemBounds;
            // CONSIDER: Add a second ExcelToolTipWindow here, and try to keep track of the one shown for the list?

            public override FormulaEdit WithFormulaEditWindow(IntPtr newFormulaEditWindow)
            {
                return new FunctionList
                {
                    FormulaEditWindow = newFormulaEditWindow,
                    FormulaPrefix = FormulaPrefix,
                    EditWindowBounds = EditWindowBounds,
                    ExcelToolTipWindow = ExcelToolTipWindow,

                    FunctionListWindow = FunctionListWindow,
                    SelectedItemText = SelectedItemText,
                    SelectedItemBounds = SelectedItemBounds,
                    FunctionListBounds = FunctionListBounds
                };
            }

            public FunctionList WithFunctionListWindow(IntPtr newFunctionListWindow)
            {
                return new FunctionList
                {
                    FormulaEditWindow = FormulaEditWindow,
                    FormulaPrefix = FormulaPrefix,
                    EditWindowBounds = EditWindowBounds,
                    ExcelToolTipWindow = ExcelToolTipWindow,

                    FunctionListWindow = newFunctionListWindow,
                    SelectedItemText = SelectedItemText,
                    SelectedItemBounds = SelectedItemBounds,
                    FunctionListBounds = FunctionListBounds
                };
            }

            public override FormulaEdit WithFormulaPrefix(string newFormulaPrefix)
            {
                return new FunctionList
                {
                    FormulaEditWindow = FormulaEditWindow,
                    FormulaPrefix = newFormulaPrefix ?? "",
                    EditWindowBounds = EditWindowBounds,
                    ExcelToolTipWindow = ExcelToolTipWindow,

                    FunctionListWindow = FunctionListWindow,
                    SelectedItemText = SelectedItemText,
                    SelectedItemBounds = SelectedItemBounds,
                    FunctionListBounds = FunctionListBounds
                };
            }

            public override FormulaEdit WithBounds(Rect newEditWindowBounds)
            {
                return new FunctionList
                {
                    FormulaEditWindow = FormulaEditWindow,
                    FormulaPrefix = FormulaPrefix,
                    EditWindowBounds = newEditWindowBounds,
                    ExcelToolTipWindow = ExcelToolTipWindow,

                    FunctionListWindow = FunctionListWindow,
                    SelectedItemText = SelectedItemText,
                    SelectedItemBounds = SelectedItemBounds,
                    FunctionListBounds = FunctionListBounds
                };
            }

            public override FormulaEdit WithToolTipWindow(IntPtr newExcelToolTipWindow)
            {
                return new FunctionList
                {
                    FormulaEditWindow = FormulaEditWindow,
                    FormulaPrefix = FormulaPrefix,
                    EditWindowBounds = EditWindowBounds,
                    ExcelToolTipWindow = newExcelToolTipWindow,

                    FunctionListWindow = FunctionListWindow,
                    SelectedItemText = SelectedItemText,
                    SelectedItemBounds = SelectedItemBounds,
                    FunctionListBounds = FunctionListBounds
                };
            }

            public virtual FunctionList WithSelectedItem(string selectedItemText, Rect selectedItemBounds, Rect listBounds)
            {
                return new FunctionList
                {
                    FormulaEditWindow = FormulaEditWindow,
                    FormulaPrefix = FormulaPrefix,
                    EditWindowBounds = EditWindowBounds,
                    ExcelToolTipWindow = ExcelToolTipWindow,

                    FunctionListWindow = FunctionListWindow,
                    FunctionListBounds = listBounds,
                    SelectedItemText = selectedItemText,
                    SelectedItemBounds = selectedItemBounds,
                };
            }

            internal FormulaEdit AsFormulaEdit()
            {
                return new FormulaEdit
                {
                    FormulaEditWindow = FormulaEditWindow,
                    FormulaPrefix = FormulaPrefix,
                    EditWindowBounds = EditWindowBounds,
                    ExcelToolTipWindow = ExcelToolTipWindow,
                };
            }
        }

        // Becomes a more general Window or Dialog to watch (for resize extension)
        public class SelectDataSource : UIState
        {
            public IntPtr SelectDataSourceWindow = IntPtr.Zero;
        }

        public override string ToString() =>
#if DEBUG
                $"{GetType().Name}{((this is Ready) ? "" : "\r\n")}{string.Join("\r\n", GetType().GetFields().Select(fld => $"\t{fld.Name}: {fld.GetValue(this)}"))}";
#else
                return base.ToString();
#endif


        // TODO: Figure out what to do with this
        public string LogString() =>
#if DEBUG
                $"{GetType().Name}{((this is Ready) ? "" : "\t")}{string.Join("\t", GetType().GetFields().Select(fld => $"\t{fld.Name}: {fld.GetValue(this)}"))}";
#else
                return ToString();
#endif


        // This is the universal update check
        // When an event knows exactly what changed (e.g. Text or SelectedItem), it need not call this
        // CONSIDER: How would this look with C# 7.0 patterns?
        public static IEnumerable<UIStateUpdate> GetUpdates(UIState oldState, UIState newState)
        {
            if (oldState is Ready)
            {
                if (newState is Ready)
                {
                    yield break;
                }
                else if (newState is FunctionList functionList)
                {
                    // We generate an intermediate state (!?)
                    FormulaEdit formulaEdit = functionList.AsFormulaEdit();
                    yield return new UIStateUpdate(oldState, formulaEdit, UIStateUpdate.UpdateType.FormulaEditStart);
                    yield return new UIStateUpdate(formulaEdit, functionList, UIStateUpdate.UpdateType.FunctionListShow);
                }
                else if (newState is FormulaEdit) // But not FunctionList
                {
                    yield return new UIStateUpdate(oldState, newState, UIStateUpdate.UpdateType.FormulaEditStart);
                }
                else if (newState is SelectDataSource)
                {
                    // Go to Ready then to new state
                    foreach (UIStateUpdate update in GetUpdates(oldState, ReadyState))
                    {
                        yield return update;
                    }

                    yield return new UIStateUpdate(ReadyState, newState, UIStateUpdate.UpdateType.SelectDataSourceShow);
                }
            }
            else if (oldState is FunctionList oldStateFl)  // and thus also FormulaEdit
            {
                if (newState is Ready)
                {
                    // We generate an intermediate state (!?)
                    FormulaEdit formulaEdit = oldStateFl.AsFormulaEdit();
                    yield return new UIStateUpdate(oldStateFl, formulaEdit, UIStateUpdate.UpdateType.FunctionListHide);
                    yield return new UIStateUpdate(formulaEdit, newState, UIStateUpdate.UpdateType.FormulaEditEnd);
                }
                else if (newState is FunctionList newStateFl)
                {
                    foreach (UIStateUpdate update in GetUpdates(oldStateFl, newStateFl))
                    {
                        yield return update;
                    }
                }
                else if (newState is FormulaEdit newStateFe) // but not FunctionList
                {
                    FormulaEdit oldStateFE = oldStateFl.AsFormulaEdit();
                    yield return new UIStateUpdate(oldStateFl, oldStateFE, UIStateUpdate.UpdateType.FunctionListHide);

                    foreach (UIStateUpdate update in GetUpdates(oldStateFE, newStateFe))
                    {
                        yield return update;
                    }
                }
                else if (newState is SelectDataSource)
                {
                    // Go to Ready then to new state
                    foreach (UIStateUpdate update in GetUpdates(oldStateFl, ReadyState))
                    {
                        yield return update;
                    }

                    yield return new UIStateUpdate(ReadyState, newState, UIStateUpdate.UpdateType.SelectDataSourceShow);
                }
            }
            else if (oldState is FormulaEdit oldStateFe)   // but not FunctionList
            {
                if (newState is Ready)
                {
                    yield return new UIStateUpdate(oldStateFe, newState, UIStateUpdate.UpdateType.FormulaEditEnd);
                }
                else if (newState is FunctionList list)
                {
                    // First process any FormulaEdit changes
                    FormulaEdit newStateFE = list.AsFormulaEdit();
                    foreach (UIStateUpdate update in GetUpdates(oldStateFe, newStateFE))
                    {
                        yield return update;
                    }

                    yield return new UIStateUpdate(newStateFE, list, UIStateUpdate.UpdateType.FunctionListShow);
                }
                else if (newState is FormulaEdit newStateFe) // but not FunctionList
                {
                    FormulaEdit oldStateFE = oldStateFe;
                    foreach (UIStateUpdate update in GetUpdates(oldStateFE, newStateFe))
                    {
                        yield return update;
                    }
                }
                else if (newState is SelectDataSource)
                {
                    // Go to Ready then to new state
                    foreach (UIStateUpdate update in GetUpdates(oldStateFe, ReadyState))
                    {
                        yield return update;
                    }

                    yield return new UIStateUpdate(ReadyState, newState, UIStateUpdate.UpdateType.SelectDataSourceShow);
                }
            }
            else if (oldState is SelectDataSource oldStateSds)
            {
                if (newState is Ready)
                {
                    yield return new UIStateUpdate(oldStateSds, newState, UIStateUpdate.UpdateType.SelectDataSourceHide);
                }
                else if (newState is SelectDataSource newStateSds)
                {
                    if (oldStateSds.SelectDataSourceWindow != newStateSds.SelectDataSourceWindow)
                    {
                        yield return new UIStateUpdate(oldStateSds, newStateSds, UIStateUpdate.UpdateType.SelectDataSourceWindowChange);
                    }
                }
                else
                {
                    // Go to Ready, then to new state
                    yield return new UIStateUpdate(oldStateSds, ReadyState, UIStateUpdate.UpdateType.SelectDataSourceHide);
                    foreach (UIStateUpdate update in GetUpdates(ReadyState, newState))
                    {
                        yield return update;
                    }
                }
            }
        }

        private static IEnumerable<UIStateUpdate> GetUpdates(FormulaEdit oldState, FormulaEdit newState)
        {
            // We generate intermediate states (!?)
            if (oldState.FormulaEditWindow != newState.FormulaEditWindow)
            {
                // Always changes together with Move ...?
                FormulaEdit tempState = oldState.WithFormulaEditWindow(newState.FormulaEditWindow);
                yield return new UIStateUpdate(oldState, tempState, UIStateUpdate.UpdateType.FormulaEditWindowChange);
                oldState = tempState;
            }
            if (oldState.EditWindowBounds != newState.EditWindowBounds)
            {
                FormulaEdit tempState = oldState.WithBounds(newState.EditWindowBounds);
                yield return new UIStateUpdate(oldState, tempState, UIStateUpdate.UpdateType.FormulaEditMove);
                oldState = tempState;
            }
            if (oldState.ExcelToolTipWindow != newState.ExcelToolTipWindow)
            {
                FormulaEdit tempState = oldState.WithToolTipWindow(newState.ExcelToolTipWindow);
                yield return new UIStateUpdate(oldState, tempState, UIStateUpdate.UpdateType.FormulaEditExcelToolTipChange);
                oldState = tempState;
            }
            if (oldState.FormulaPrefix != newState.FormulaPrefix)
            {
                yield return new UIStateUpdate(oldState, newState, UIStateUpdate.UpdateType.FormulaEditTextChange);
            }
        }

        private static IEnumerable<UIStateUpdate> GetUpdates(FunctionList oldState, FunctionList newState)
        {
            // We generate intermediate states (!?)
            if (oldState.FormulaEditWindow != newState.FormulaEditWindow)
            {
                // Always changes together with Move ...?
                FormulaEdit tempState = oldState.WithFormulaEditWindow(newState.FormulaEditWindow);
                yield return new UIStateUpdate(oldState, tempState, UIStateUpdate.UpdateType.FormulaEditWindowChange);
                oldState = (FunctionList)tempState;
            }
            if (oldState.FunctionListWindow != newState.FunctionListWindow)
            {
                Debug.Print(">>>>> Unexpected FunctionListWindowChange");  // Should never change???
                FunctionList tempState = oldState.WithFunctionListWindow(newState.FunctionListWindow);
                yield return new UIStateUpdate(oldState, tempState, UIStateUpdate.UpdateType.FunctionListWindowChange);
                oldState = tempState;
            }
            if (oldState.EditWindowBounds != newState.EditWindowBounds)
            {
                FormulaEdit tempState = oldState.WithBounds(newState.EditWindowBounds);
                yield return new UIStateUpdate(oldState, tempState, UIStateUpdate.UpdateType.FormulaEditMove);
                oldState = (FunctionList)tempState;
            }
            if (oldState.ExcelToolTipWindow != newState.ExcelToolTipWindow)
            {
                FormulaEdit tempState = oldState.WithToolTipWindow(newState.ExcelToolTipWindow);
                yield return new UIStateUpdate(oldState, tempState, UIStateUpdate.UpdateType.FormulaEditExcelToolTipChange);
                oldState = (FunctionList)tempState;
            }
            if (oldState.FormulaPrefix != newState.FormulaPrefix)
            {
                FormulaEdit tempState = oldState.WithFormulaPrefix(newState.FormulaPrefix);
                yield return new UIStateUpdate(oldState, tempState, UIStateUpdate.UpdateType.FormulaEditTextChange);
                oldState = (FunctionList)tempState;
            }
            if (oldState.SelectedItemText != newState.SelectedItemText ||
                oldState.SelectedItemBounds != newState.SelectedItemBounds ||
                oldState.FunctionListBounds != newState.FunctionListBounds)
            {
                yield return new UIStateUpdate(oldState, newState, UIStateUpdate.UpdateType.FunctionListSelectedItemChange);
            }
        }
    }

    internal class UIStateUpdate : EventArgs
    {
        // We want to order and nest the updates to make them easy to respond to.
        // This means we have XXXStart, then stuff on the inside, then XXXEnd, always with correct nesting
        public enum UpdateType
        {
            FormulaEditStart,
            // These three updates can happen while FunctionList is shown
            FormulaEditMove,    // Includes moving between in-cell editing and the formula text box.
            FormulaEditWindowChange, // Includes moving between in-cell editing and the formula text box.
            FormulaEditTextChange,
            FormulaEditExcelToolTipChange,

            FunctionListShow,
            FunctionListMove,
            FunctionListSelectedItemChange,
            FunctionListWindowChange,
            FunctionListHide,

            FormulaEditEnd,

            SelectDataSourceShow,
            SelectDataSourceWindowChange,
            SelectDataSourceHide
        }
        public UIState OldState { get; }
        public UIState NewState { get; }
        public UpdateType Update { get; }
        public bool IsEnabled { get; private set; }  // Should this update be raised on the main thread - allows preview event to filter out some events

        public UIStateUpdate(UIState oldState, UIState newState, UpdateType update)
        {
            OldState = oldState;
            NewState = newState;
            Update = update;
            IsEnabled = false;
        }

        // Call this to allow the update event (on the main thread) to be raised
        public void EnableUpdateEvent() => IsEnabled = true;

        public override string ToString() => $"{Update}: {OldState} -> {NewState}";

        public string LogString() => $"({Update.ToString()}): [{OldState.LogString()}] -> [{NewState.LogString()}]";
    }
}
