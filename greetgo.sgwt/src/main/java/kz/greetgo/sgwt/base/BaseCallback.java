package kz.greetgo.sgwt.base;

import kz.greetgo.gwtshare.base.ClientException;
import kz.greetgo.gwtshare.base.SgwtException;

import com.google.gwt.user.client.Window;
import com.google.gwt.user.client.rpc.AsyncCallback;
import com.google.gwt.user.client.rpc.StatusCodeException;
import com.smartgwt.client.util.SC;

public abstract class BaseCallback<R> implements AsyncCallback<R> {
  
  @Override
  public void onFailure(Throwable caught) {
    
    if (caught instanceof StatusCodeException
        && ((StatusCodeException)caught).getStatusCode() == 401) {
      Window.Location.reload();
      return;
    }
    
    if (caught instanceof ClientException) {
      SC.say(caught.getMessage());
      return;
    }
    
    if (Window.Location.getParameter("hideerrors") == null) {
      boolean isShow = caught instanceof SgwtException
          || Window.Location.getParameter("showerrors") != null;
      if (isShow) SC.warn(caught.getMessage());
    }
    
    caught.printStackTrace();
  }
}
