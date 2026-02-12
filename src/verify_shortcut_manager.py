import os
import shutil
from shortcut_manager import ShortcutManager

def test_shortcut_manager():
    print("Testing ShortcutManager...")
    
    # 1. Setup dummy environment
    base_dir = "test_shortcuts"
    if os.path.exists(base_dir):
        shutil.rmtree(base_dir)
    os.makedirs(base_dir)
    
    shortcuts_dir = os.path.join(base_dir, "shortcuts")
    os.makedirs(shortcuts_dir)
    
    # Create valid app definition files
    with open(os.path.join(shortcuts_dir, "app1.tsv"), "w", encoding="utf-8") as f:
        f.write("Ctrl+C\tCopy\n")
        f.write("Ctrl+V\tPaste\n")
        
    with open(os.path.join(shortcuts_dir, "app2.tsv"), "w", encoding="utf-8") as f:
        f.write("Ctrl+C\tMeow\n") # Duplicate key, diff desc
        f.write("Ctrl+X\tCut\n")
        
    manager_path = os.path.join(base_dir, "shortcuts.tsv")
    manager = ShortcutManager(manager_path)
    
    # 2. Test App Definition Loading
    print("\n--- Testing App Definition Loading ---")
    files = manager.get_app_definitions_list()
    print(f"App files found: {files}")
    assert "app1.tsv" in files
    assert "app2.tsv" in files
    
    # Load separate
    desc1 = manager.load_app_definitions(["app1.tsv"])
    print(f"App1 Desc: {desc1.get('Ctrl+C')}") 
    assert desc1.get("Ctrl+C") == "Copy"
    
    # Load combined
    desc_combined = manager.load_app_definitions(["app1.tsv", "app2.tsv"])
    print(f"Combined Ctrl+C: {desc_combined.get('Ctrl+C')}")
    # Expect "Copy、Meow" (order depends on list order, but set logic might vary? No, implementation uses list traversal)
    # The implementation:
    # app1: desc["Ctrl+C"] = ["Copy"]
    # app2: desc["Ctrl+C"] = ["Copy", "Meow"]
    # join -> "Copy、Meow"
    assert "Copy" in desc_combined.get("Ctrl+C")
    assert "Meow" in desc_combined.get("Ctrl+C")
    
    print(f"Combined Ctrl+X: {desc_combined.get('Ctrl+X')}")
    assert desc_combined.get("Ctrl+X") == "Cut"
    
    # 3. Test Record/Play Logic
    print("\n--- Testing Record/Play Logic ---")
    # Default is enabled for modifiers+key? 
    # Ctrl+C should be enabled by default.
    assert manager.is_allowed_record("Ctrl+C")
    assert manager.is_allowed_playback("Ctrl+C")
    
    # Update
    manager.update_shortcut("Ctrl+C", False, True)
    assert not manager.is_allowed_record("Ctrl+C")
    assert manager.is_allowed_playback("Ctrl+C")
    
    manager.save()
    
    # Reload
    manager2 = ShortcutManager(manager_path)
    assert not manager2.is_allowed_record("Ctrl+C")
    assert manager2.is_allowed_playback("Ctrl+C")
    
    print("Verification Passed!")
    
    # Cleanup
    # shutil.rmtree(base_dir)

if __name__ == "__main__":
    test_shortcut_manager()
