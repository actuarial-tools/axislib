'''shellcmd - simple invocation of shell commands from Python'''
class shellcmd:
  def __init__(self, cmd=None):
        self.cmd = cmd        
              
  def shell_call(self, cmd, *args, **kwds):
        if args or kwds:
            cmd = cmd.format(*args, **kwds)
        return subprocess.call(cmd, shell=True)

    def check_shell_call(self, cmd, *args, **kwds):
        if args or kwds:
            cmd = cmd.format(*args, **kwds)
        return subprocess.check_call(cmd, shell=True)

    def check_shell_output(self, cmd, *args, **kwds):
        if args or kwds:
            cmd = cmd.format(*args, **kwds)
        return subprocess.check_output(cmd, shell=True)
    
import shellcmd
return_code = shellcmd.shell_call('ls -l {}', dirname)
listing = shellcmd.check_shell_output('ls -l {}', dirname)
