package com.automationanywhere.botcommand.samples.commands.basic;

import com.automationanywhere.bot.service.GlobalSessionContext;
import com.automationanywhere.botcommand.BotCommand;
import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.commandsdk.i18n.Messages;
import com.automationanywhere.commandsdk.i18n.MessagesFactory;
import java.lang.Boolean;
import java.lang.ClassCastException;
import java.lang.Deprecated;
import java.lang.Double;
import java.lang.NullPointerException;
import java.lang.Object;
import java.lang.String;
import java.lang.Throwable;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;
import java.util.Optional;
import java.util.stream.Collectors;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

public final class DataFormatCommand implements BotCommand {
  private static final Logger logger = LogManager.getLogger(DataFormatCommand.class);

  private static final Messages MESSAGES_GENERIC = MessagesFactory.getMessages("com.automationanywhere.commandsdk.generic.messages");

  @Deprecated
  public Optional<Value> execute(Map<String, Value> parameters, Map<String, Object> sessionMap) {
    return execute(null, parameters, sessionMap);
  }

  public Optional<Value> execute(GlobalSessionContext globalSessionContext,
      Map<String, Value> parameters, Map<String, Object> sessionMap) {
    logger.traceEntry(() -> parameters != null ? parameters.entrySet().stream().filter(en -> !Arrays.asList( new String[] {}).contains(en.getKey()) && en.getValue() != null).collect(Collectors.toMap(Map.Entry::getKey, Map.Entry::getValue)).toString() : null, ()-> sessionMap != null ?sessionMap.toString() : null);
    DataFormat command = new DataFormat();
    HashMap<String, Object> convertedParameters = new HashMap<String, Object>();
    if(parameters.containsKey("file") && parameters.get("file") != null && parameters.get("file").get() != null) {
      convertedParameters.put("file", parameters.get("file").get());
      if(convertedParameters.get("file") !=null && !(convertedParameters.get("file") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","file", "String", parameters.get("file").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("file") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","file"));
    }

    if(parameters.containsKey("getSheetBy") && parameters.get("getSheetBy") != null && parameters.get("getSheetBy").get() != null) {
      convertedParameters.put("getSheetBy", parameters.get("getSheetBy").get());
      if(convertedParameters.get("getSheetBy") !=null && !(convertedParameters.get("getSheetBy") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","getSheetBy", "String", parameters.get("getSheetBy").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("getSheetBy") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","getSheetBy"));
    }
    if(convertedParameters.get("getSheetBy") != null) {
      switch((String)convertedParameters.get("getSheetBy")) {
        case "name" : {
          if(parameters.containsKey("sheetName") && parameters.get("sheetName") != null && parameters.get("sheetName").get() != null) {
            convertedParameters.put("sheetName", parameters.get("sheetName").get());
            if(convertedParameters.get("sheetName") !=null && !(convertedParameters.get("sheetName") instanceof String)) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","sheetName", "String", parameters.get("sheetName").get().getClass().getSimpleName()));
            }
          }
          if(convertedParameters.get("sheetName") == null) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","sheetName"));
          }


        } break;
        case "index" : {
          if(parameters.containsKey("sheetIndex") && parameters.get("sheetIndex") != null && parameters.get("sheetIndex").get() != null) {
            convertedParameters.put("sheetIndex", parameters.get("sheetIndex").get());
            if(convertedParameters.get("sheetIndex") !=null && !(convertedParameters.get("sheetIndex") instanceof Double)) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","sheetIndex", "Double", parameters.get("sheetIndex").get().getClass().getSimpleName()));
            }
          }
          if(convertedParameters.get("sheetIndex") == null) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","sheetIndex"));
          }
          if(convertedParameters.containsKey("sheetIndex")) {
            try {
              if(convertedParameters.get("sheetIndex") != null && !((double)convertedParameters.get("sheetIndex") >= 0)) {
                throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.GreaterThanEqualTo","sheetIndex", "0"));
              }
            }
            catch(ClassCastException e) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","sheetIndex", "Number", convertedParameters.get("sheetIndex").getClass().getSimpleName()));
            }
            catch(NullPointerException e) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","sheetIndex"));
            }

          }

        } break;
        default : throw new BotCommandException(MESSAGES_GENERIC.getString("generic.InvalidOption","getSheetBy"));
      }
    }

    if(parameters.containsKey("range") && parameters.get("range") != null && parameters.get("range").get() != null) {
      convertedParameters.put("range", parameters.get("range").get());
      if(convertedParameters.get("range") !=null && !(convertedParameters.get("range") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","range", "String", parameters.get("range").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("range") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","range"));
    }

    if(parameters.containsKey("formato") && parameters.get("formato") != null && parameters.get("formato").get() != null) {
      convertedParameters.put("formato", parameters.get("formato").get());
      if(convertedParameters.get("formato") !=null && !(convertedParameters.get("formato") instanceof Boolean)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","formato", "Boolean", parameters.get("formato").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("formato") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","formato"));
    }
    if(convertedParameters.get("formato") != null && (Boolean)convertedParameters.get("formato")) {
      if(parameters.containsKey("formato_value") && parameters.get("formato_value") != null && parameters.get("formato_value").get() != null) {
        convertedParameters.put("formato_value", parameters.get("formato_value").get());
        if(convertedParameters.get("formato_value") !=null && !(convertedParameters.get("formato_value") instanceof String)) {
          throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","formato_value", "String", parameters.get("formato_value").get().getClass().getSimpleName()));
        }
      }
      if(convertedParameters.get("formato_value") == null) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","formato_value"));
      }
      if(convertedParameters.get("formato_value") != null) {
        switch((String)convertedParameters.get("formato_value")) {
          case "0" : {

          } break;
          case "2" : {

          } break;
          case "165" : {

          } break;
          case "44" : {

          } break;
          case "14" : {

          } break;
          case "166" : {

          } break;
          case "167" : {

          } break;
          case "10" : {

          } break;
          case "12" : {

          } break;
          case "11" : {

          } break;
          case "49" : {

          } break;
          case "custom" : {
            if(parameters.containsKey("custom_value") && parameters.get("custom_value") != null && parameters.get("custom_value").get() != null) {
              convertedParameters.put("custom_value", parameters.get("custom_value").get());
              if(convertedParameters.get("custom_value") !=null && !(convertedParameters.get("custom_value") instanceof String)) {
                throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","custom_value", "String", parameters.get("custom_value").get().getClass().getSimpleName()));
              }
            }
            if(convertedParameters.get("custom_value") == null) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","custom_value"));
            }


          } break;
          default : throw new BotCommandException(MESSAGES_GENERIC.getString("generic.InvalidOption","formato_value"));
        }
      }

    }

    if(parameters.containsKey("alinhamento") && parameters.get("alinhamento") != null && parameters.get("alinhamento").get() != null) {
      convertedParameters.put("alinhamento", parameters.get("alinhamento").get());
      if(convertedParameters.get("alinhamento") !=null && !(convertedParameters.get("alinhamento") instanceof Boolean)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","alinhamento", "Boolean", parameters.get("alinhamento").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("alinhamento") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","alinhamento"));
    }
    if(convertedParameters.get("alinhamento") != null && (Boolean)convertedParameters.get("alinhamento")) {
      if(parameters.containsKey("alinhamento_value") && parameters.get("alinhamento_value") != null && parameters.get("alinhamento_value").get() != null) {
        convertedParameters.put("alinhamento_value", parameters.get("alinhamento_value").get());
        if(convertedParameters.get("alinhamento_value") !=null && !(convertedParameters.get("alinhamento_value") instanceof String)) {
          throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","alinhamento_value", "String", parameters.get("alinhamento_value").get().getClass().getSimpleName()));
        }
      }
      if(convertedParameters.get("alinhamento_value") == null) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","alinhamento_value"));
      }
      if(convertedParameters.get("alinhamento_value") != null) {
        switch((String)convertedParameters.get("alinhamento_value")) {
          case "LEFT" : {

          } break;
          case "CENTER" : {

          } break;
          case "RIGHT" : {

          } break;
          default : throw new BotCommandException(MESSAGES_GENERIC.getString("generic.InvalidOption","alinhamento_value"));
        }
      }

    }

    if(parameters.containsKey("alinhamentoVertical") && parameters.get("alinhamentoVertical") != null && parameters.get("alinhamentoVertical").get() != null) {
      convertedParameters.put("alinhamentoVertical", parameters.get("alinhamentoVertical").get());
      if(convertedParameters.get("alinhamentoVertical") !=null && !(convertedParameters.get("alinhamentoVertical") instanceof Boolean)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","alinhamentoVertical", "Boolean", parameters.get("alinhamentoVertical").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("alinhamentoVertical") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","alinhamentoVertical"));
    }
    if(convertedParameters.get("alinhamentoVertical") != null && (Boolean)convertedParameters.get("alinhamentoVertical")) {
      if(parameters.containsKey("alinhamentoVertical_value") && parameters.get("alinhamentoVertical_value") != null && parameters.get("alinhamentoVertical_value").get() != null) {
        convertedParameters.put("alinhamentoVertical_value", parameters.get("alinhamentoVertical_value").get());
        if(convertedParameters.get("alinhamentoVertical_value") !=null && !(convertedParameters.get("alinhamentoVertical_value") instanceof String)) {
          throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","alinhamentoVertical_value", "String", parameters.get("alinhamentoVertical_value").get().getClass().getSimpleName()));
        }
      }
      if(convertedParameters.get("alinhamentoVertical_value") == null) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","alinhamentoVertical_value"));
      }
      if(convertedParameters.get("alinhamentoVertical_value") != null) {
        switch((String)convertedParameters.get("alinhamentoVertical_value")) {
          case "TOP" : {

          } break;
          case "CENTER" : {

          } break;
          case "BOTTOM" : {

          } break;
          default : throw new BotCommandException(MESSAGES_GENERIC.getString("generic.InvalidOption","alinhamentoVertical_value"));
        }
      }

    }

    if(parameters.containsKey("bloquear") && parameters.get("bloquear") != null && parameters.get("bloquear").get() != null) {
      convertedParameters.put("bloquear", parameters.get("bloquear").get());
      if(convertedParameters.get("bloquear") !=null && !(convertedParameters.get("bloquear") instanceof Boolean)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","bloquear", "Boolean", parameters.get("bloquear").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("bloquear") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","bloquear"));
    }
    if(convertedParameters.get("bloquear") != null && (Boolean)convertedParameters.get("bloquear")) {
      if(parameters.containsKey("bloquear_value") && parameters.get("bloquear_value") != null && parameters.get("bloquear_value").get() != null) {
        convertedParameters.put("bloquear_value", parameters.get("bloquear_value").get());
        if(convertedParameters.get("bloquear_value") !=null && !(convertedParameters.get("bloquear_value") instanceof Boolean)) {
          throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","bloquear_value", "Boolean", parameters.get("bloquear_value").get().getClass().getSimpleName()));
        }
      }
      if(convertedParameters.get("bloquear_value") == null) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","bloquear_value"));
      }

    }

    try {
      command.action((String)convertedParameters.get("file"),(String)convertedParameters.get("getSheetBy"),(String)convertedParameters.get("sheetName"),(Double)convertedParameters.get("sheetIndex"),(String)convertedParameters.get("range"),(Boolean)convertedParameters.get("formato"),(String)convertedParameters.get("formato_value"),(String)convertedParameters.get("custom_value"),(Boolean)convertedParameters.get("alinhamento"),(String)convertedParameters.get("alinhamento_value"),(Boolean)convertedParameters.get("alinhamentoVertical"),(String)convertedParameters.get("alinhamentoVertical_value"),(Boolean)convertedParameters.get("bloquear"),(Boolean)convertedParameters.get("bloquear_value"));Optional<Value> result = Optional.empty();
      return logger.traceExit(result);
    }
    catch (ClassCastException e) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.IllegalParameters","action"));
    }
    catch (BotCommandException e) {
      logger.fatal(e.getMessage(),e);
      throw e;
    }
    catch (Throwable e) {
      logger.fatal(e.getMessage(),e);
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.NotBotCommandException",e.getMessage()),e);
    }
  }
}
