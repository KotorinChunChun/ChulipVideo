import argparse
import sys
import os

# Add current directory to sys.path to ensure imports work
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def main():
    parser = argparse.ArgumentParser(description="Launcher for Video Clip Creation Tools")
    parser.add_argument('mode', nargs='?', choices=['chunap', 'capchun', 'chulip'], default='chulip',
                        help='Application mode to launch (default: chulip)')
    
    args = parser.parse_args()
    
    if args.mode == 'chunap':
        import ChunapTool
        ChunapTool.main()
    elif args.mode == 'capchun':
        import CapchunScreen
        CapchunScreen.main()
    elif args.mode == 'chulip':
        import ChulipVideo
        ChulipVideo.main()

if __name__ == "__main__":
    main()
