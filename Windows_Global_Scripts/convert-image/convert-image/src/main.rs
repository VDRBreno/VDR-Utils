use std::env;
use std::io;
use std::process::Command;
use std::path::Path;

struct CustomError {
    error_ocurred: bool,
    message: String
}

impl CustomError {
    fn handle_error(&mut self, message: String) {
        self.error_ocurred = true;
        self.message = message;
    }
}

fn get_index_of_first_match(string: &String, chars_to_match: Vec<char>) -> usize {
    let mut char_index: usize = 0;
    'out: for char in string.chars() {
        for char_to_match in &chars_to_match {
            if char == *char_to_match {
                break 'out;
            }
        }
        char_index += 1;
    }
    return char_index;
}

fn get_only_filename(filename: &str) -> String {
    let mut filename_temp_splitted = split_in_characters(String::from(filename));
    filename_temp_splitted.reverse();
    let filename_temp = filename_temp_splitted.join("");
    let dot_index = get_index_of_first_match(&filename_temp, Vec::from(['.']));
    return String::from(&filename[0..(filename_temp.len()-1)-dot_index]);
}

fn split_in_characters(string: String) -> Vec<String> {
    let mut chars: Vec<String> = Vec::new();
    for char in string.chars() {
        chars.insert(chars.len(), char.to_string());
    }
    return chars;
}

fn remove_special_chars(string: String) -> String {
    let temp = string.replace("\r\n", "");
    return temp;
}

fn get_filename_from_arg(string: String) -> String {
    let mut path_temp_splitted = split_in_characters(string.clone());
    path_temp_splitted.reverse();
    let path_temp = path_temp_splitted.join("");
    let bar_index = get_index_of_first_match(&path_temp, Vec::from(['\\', '/']));
    return String::from(&string[string.len()-bar_index..string.len()]);
}

fn get_filetype_by_filename(filename: &String) -> String {
    let dot_index = get_index_of_first_match(&filename, Vec::from(['.']));
    return String::from(&filename[dot_index+1..filename.len()]);
}

fn main() {
    let mut error_handler = CustomError {
        error_ocurred: false,
        message: String::new()
    };

    let mut await_response_buffer = String::new();
    
    handle(&mut error_handler);

    if error_handler.error_ocurred == true {
        println!("\nError: {}", error_handler.message);
    }
    println!("\nProgram finished press any key...");
    io::stdin()
        .read_line(&mut await_response_buffer)
        .expect("Error to await_response_buffer");
}

fn handle(error_handler: &mut CustomError) {
    let mut args: Vec<String> = env::args().collect();
    
    // first argument is a path of script
    args.remove(0);

    if args.len() < 1 {
        error_handler.handle_error(String::from("Error: Expecting at least 1 argument, received: 0"));
        return;
    }

    let mut filename = get_filename_from_arg(args[0].clone());
    let path_not_parsed = if filename.len()!=args[0].len()
        { String::from(&args[0][0..(args[0].len()-filename.len()-1)]) }
        else { String::from(".\\") };
    let path = Path::new(&path_not_parsed).to_string_lossy().into_owned();

    let mut new_file_type = String::new();
    if args.len() > 1 {
        new_file_type = args[1].clone();
    } else {
        println!("Convert to: (Ex: png, jpg)");
        io::stdin()
            .read_line(&mut new_file_type)
            .expect("Error to read new_file_type");
    }

    if new_file_type==get_filetype_by_filename(&filename) {
        error_handler.handle_error(String::from("The type choose is equal to original..."));
        return;
    }

    filename = remove_special_chars(filename);
    let only_filename = remove_special_chars(get_only_filename(filename.as_str()));
    let output_filename = format!("{}.{}", only_filename, new_file_type);
    let fixed_filename = Path::new(&format!("{}\\{}", path, filename)).to_string_lossy().into_owned();
    let fixed_output_file = Path::new(&format!("{}\\{}", path, output_filename)).to_string_lossy().into_owned();

    let command_argument = format!("ffmpeg -i \"{}\" \"{}\"", fixed_filename, fixed_output_file);

    let mut binding = Command::new("powershell");
    let command = binding.arg(command_argument);
    let command_status = command.spawn().unwrap().wait().expect("Error while trying to run command");

    if !command_status.success() {
        error_handler.handle_error(String::from("Unable to execute the command"))
    }
}