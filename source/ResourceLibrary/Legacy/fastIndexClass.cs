
using System;
using System.Runtime.Serialization;

namespace Contensive.Addons.ResourceLibrary.Controllers {
    public class FastIndexClass {
        private const int KeyPointerArrayChunk = 1000;

        [Serializable]
        public class storageClass {
            public int ArraySize;
            public int ArrayCount;
            public bool ArrayDirty;
            public string[] UcaseKeyArray;
            public string[] PointerArray;
            public int ArrayPointer;
        }

        private storageClass store = new storageClass();

        private int GetArrayPointer(string Key) {
            int ArrayPointer = -1;
            try {
                if (store.ArrayDirty) {
                    Sort();
                }
                ArrayPointer = -1;
                if (store.ArrayCount > 0) {
                    string UcaseTargetKey = Key.ToUpper().Replace("\r\n", "");
                    int LowGuess = -1;
                    int HighGuess = store.ArrayCount - 1;
                    while ((HighGuess - LowGuess) > 1) {
                        int PointerGuess = (HighGuess + LowGuess) / 2;
                        if (UcaseTargetKey == store.UcaseKeyArray[PointerGuess]) {
                            HighGuess = PointerGuess;
                            break;
                        } else if (string.CompareOrdinal(UcaseTargetKey, store.UcaseKeyArray[PointerGuess]) < 0) {
                            HighGuess = PointerGuess;
                        } else {
                            LowGuess = PointerGuess;
                        }
                    }
                    if (UcaseTargetKey == store.UcaseKeyArray[HighGuess]) {
                        ArrayPointer = HighGuess;
                    }
                }
            } catch (Exception ex) {
                throw new indexException("getArrayPointer error", ex);
            }
            return ArrayPointer;
        }

        public int GetPointer(string key) {
            return getPtr(key);
        }

        public int getPtr(string Key) {
            int returnKey = -1;
            try {
                string UcaseKey = Key.ToUpper().Replace("\r\n", "");
                store.ArrayPointer = GetArrayPointer(Key);
                if (store.ArrayPointer > -1) {
                    bool MatchFound = true;
                    while (MatchFound) {
                        store.ArrayPointer = store.ArrayPointer - 1;
                        if (store.ArrayPointer < 0) {
                            MatchFound = false;
                        } else {
                            MatchFound = (store.UcaseKeyArray[store.ArrayPointer] == UcaseKey);
                        }
                    }
                    store.ArrayPointer = store.ArrayPointer + 1;
                    returnKey = 0;
                    if (int.TryParse(store.PointerArray[store.ArrayPointer], out _)) {
                        returnKey = int.Parse(store.PointerArray[store.ArrayPointer]);
                    }
                }
            } catch (Exception ex) {
                throw new indexException("GetPointer error", ex);
            }
            return returnKey;
        }

        public void SetPointer(string key, int pointer) {
            setPtr(key, pointer);
        }

        public void setPtr(string Key, int Pointer) {
            try {
                string keyToSave = Key.ToUpper().Replace("\r\n", "");
                if (store.ArrayCount >= store.ArraySize) {
                    store.ArraySize = store.ArraySize + KeyPointerArrayChunk;
                    Array.Resize(ref store.PointerArray, store.ArraySize + 1);
                    Array.Resize(ref store.UcaseKeyArray, store.ArraySize + 1);
                }
                store.ArrayPointer = store.ArrayCount;
                store.ArrayCount = store.ArrayCount + 1;
                store.UcaseKeyArray[store.ArrayPointer] = keyToSave;
                store.PointerArray[store.ArrayPointer] = Pointer.ToString();
                store.ArrayDirty = true;
            } catch (Exception ex) {
                throw new indexException("SetPointer error", ex);
            }
        }

        public int GetNextPointerMatch(string key) {
            return getNextPtrMatch(key);
        }

        public int getNextPtrMatch(string Key) {
            int nextPointerMatch = -1;
            try {
                if (store.ArrayPointer < (store.ArrayCount - 1)) {
                    store.ArrayPointer = store.ArrayPointer + 1;
                    string UcaseKey = Key.ToUpper();
                    if (store.UcaseKeyArray[store.ArrayPointer] == UcaseKey) {
                        if (int.TryParse(store.PointerArray[store.ArrayPointer], out _)) {
                            nextPointerMatch = int.Parse(store.PointerArray[store.ArrayPointer]);
                        }
                    } else {
                        store.ArrayPointer = store.ArrayPointer - 1;
                    }
                }
            } catch (Exception ex) {
                throw new indexException("GetNextPointerMatch error", ex);
            }
            return nextPointerMatch;
        }

        public int GetFirstPointer() {
            return getFirstPtr();
        }

        public int getFirstPtr() {
            int firstPointer = -1;
            try {
                if (store.ArrayDirty) {
                    Sort();
                }
                if (store.ArrayCount > 0) {
                    store.ArrayPointer = 0;
                    firstPointer = 0;
                    if (int.TryParse(store.PointerArray[store.ArrayPointer], out _)) {
                        firstPointer = int.Parse(store.PointerArray[store.ArrayPointer]);
                    }
                }
            } catch (Exception ex) {
                throw new indexException("GetFirstPointer error", ex);
            }
            return firstPointer;
        }

        public int GetNextPointer() {
            return getNextPtr();
        }

        public int getNextPtr() {
            int nextPointer = -1;
            try {
                if (store.ArrayDirty) {
                    Sort();
                }
                if ((store.ArrayPointer + 1) < store.ArrayCount) {
                    store.ArrayPointer = store.ArrayPointer + 1;
                    nextPointer = 0;
                    if (int.TryParse(store.PointerArray[store.ArrayPointer], out _)) {
                        nextPointer = int.Parse(store.PointerArray[store.ArrayPointer]);
                    }
                }
            } catch (Exception ex) {
                throw new indexException("GetPointer error", ex);
            }
            return nextPointer;
        }

        private void BubbleSort() {
            try {
                if (store.ArrayCount > 1) {
                    int PointerDelta = 1;
                    int MaxPointer = store.ArrayCount - 2;
                    for (int SlowPointer = MaxPointer; SlowPointer >= 0; SlowPointer--) {
                        bool CleanPass = true;
                        for (int FastPointer = MaxPointer; FastPointer >= (MaxPointer - SlowPointer); FastPointer--) {
                            if (string.CompareOrdinal(store.UcaseKeyArray[FastPointer], store.UcaseKeyArray[FastPointer + PointerDelta]) > 0) {
                                string TempUcaseKey = store.UcaseKeyArray[FastPointer + PointerDelta];
                                string tempPtrString = store.PointerArray[FastPointer + PointerDelta];
                                store.UcaseKeyArray[FastPointer + PointerDelta] = store.UcaseKeyArray[FastPointer];
                                store.PointerArray[FastPointer + PointerDelta] = store.PointerArray[FastPointer];
                                store.UcaseKeyArray[FastPointer] = TempUcaseKey;
                                store.PointerArray[FastPointer] = tempPtrString;
                                CleanPass = false;
                            }
                        }
                        if (CleanPass) {
                            break;
                        }
                    }
                }
                store.ArrayDirty = false;
            } catch (Exception ex) {
                throw new indexException("BubbleSort error", ex);
            }
        }

        private void QuickSort() {
            try {
                if (store.ArrayCount >= 2) {
                    QuickSort_Segment(store.UcaseKeyArray, store.PointerArray, 0, store.ArrayCount - 1);
                }
            } catch (Exception ex) {
                throw new indexException("QuickSort error", ex);
            }
        }

        private void QuickSort_Segment(string[] C, string[] P, int First, int Last) {
            try {
                int Low = First;
                int High = Last;
                string MidValue = C[(First + Last) / 2];
                do {
                    while (string.CompareOrdinal(C[Low], MidValue) < 0) {
                        Low = Low + 1;
                    }
                    while (string.CompareOrdinal(C[High], MidValue) > 0) {
                        High = High - 1;
                    }
                    if (Low <= High) {
                        string TC = C[Low];
                        string TP = P[Low];
                        C[Low] = C[High];
                        P[Low] = P[High];
                        C[High] = TC;
                        P[High] = TP;
                        Low = Low + 1;
                        High = High - 1;
                    }
                } while (Low <= High);
                if (First < High) {
                    QuickSort_Segment(C, P, First, High);
                }
                if (Low < Last) {
                    QuickSort_Segment(C, P, Low, Last);
                }
            } catch (Exception ex) {
                throw new indexException("QuickSort_Segment error", ex);
            }
        }

        private void Sort() {
            try {
                QuickSort();
                store.ArrayDirty = false;
            } catch (Exception ex) {
                throw new indexException("Sort error", ex);
            }
        }
    }

    public class indexException : System.Exception, System.Runtime.Serialization.ISerializable {
        public indexException() : base() { }
        public indexException(string message) : base(message) { }
        public indexException(string message, Exception inner) : base(message, inner) { }
        protected indexException(System.Runtime.Serialization.SerializationInfo info, System.Runtime.Serialization.StreamingContext context) : base(info, context) { }
    }
}
